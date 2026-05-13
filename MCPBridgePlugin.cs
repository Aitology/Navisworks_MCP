// =============================================================================
//  Navisworks MCP Bridge — full implementation (35 routes)
// =============================================================================
//  >>> VERSION MARKER: FULL-V2 <<<
//
//  In-process Navisworks plugin that hosts a tiny HTTP server on
//  http://localhost:8765/. The Python MCP server (../server_http.py) POSTs
//  JSON to it; we marshal each call onto the Navisworks UI thread, hit the
//  API, and return JSON.
//
//  Routes
//  ──────
//  Health/Document  (3)  /health, /document/info, /document/statistics
//  Selection        (6)  /selection/{get,select-all,deselect-all,invert,
//                                   from-set,by-name}
//  Sets             (6)  /sets/{list,selection/create,search/create,delete,
//                              items,search/run}
//  Find/Properties  (3)  /elements/{find,find-by-name,properties}
//  Overrides        (6)  /elements/{color,transparency,hide,unhide-all,
//                                   isolate,reset-overrides}
//  Clash            (5)  /clash/{list,create,run,results,delete}
//  Viewpoints       (5)  /viewpoints/{list,save,goto,delete,current}
//
//  Lifecycle
//  ─────────
//  Listener auto-starts in the static constructor when Navisworks loads the
//  DLL. The ribbon button is a manual restart / status dialog — useful as a
//  fallback but not normally needed.
//
//  Threading
//  ─────────
//  HttpListener delivers requests on worker threads, but the Navisworks API
//  is single-threaded (UI). Every handler runs through RunOnUi(), which uses
//  the SynchronizationContext captured at plugin load and Send()s synchronously
//  so the HTTP response naturally serializes the result.
//
//  Logs
//  ────
//  %TEMP%\navisworks_mcp_addin.log
// =============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Clash;
using Autodesk.Navisworks.Api.Plugins;

using NwColor = Autodesk.Navisworks.Api.Color;
using NavisApp = Autodesk.Navisworks.Api.Application;

namespace NavisworksMcpAddin
{
    [Plugin("MCPBridge", "ADSK",
            DisplayName = "MCP Bridge",
            ToolTip = "HTTP bridge for the Model Context Protocol server")]
    public class MCPBridgePlugin : AddInPlugin
    {
        // ── Configuration ───────────────────────────────────────────────────
        private const int Port = 8765;
        private const string Version = "2.1";
        private static readonly string LogPath =
            Path.Combine(Path.GetTempPath(), "navisworks_mcp_addin.log");

        // ── State ───────────────────────────────────────────────────────────
        private static HttpListener _listener;
        private static Thread _listenerThread;
        private static SynchronizationContext _uiContext;
        private static readonly object _startLock = new object();

        // ─────────────────────────────────────────────────────────────────────
        //  Plugin lifecycle
        // ─────────────────────────────────────────────────────────────────────

        // Static ctor runs on the UI thread when Navisworks first touches the
        // type during plugin discovery. We capture the SynchronizationContext
        // and start the listener so the bridge is live before the user does
        // anything.
        static MCPBridgePlugin()
        {
            try
            {
                Log($"Plugin DLL loaded. Version: FULL-V2 (bridge {Version})");
                _uiContext = SynchronizationContext.Current;
                if (_uiContext == null)
                {
                    Log("WARN: SynchronizationContext.Current was null in static ctor.");
                }
                TryStartListener();
            }
            catch (Exception ex)
            {
                Log("Static ctor failed: " + ex);
            }
        }

        // Manual restart / status check. Capture UI context here too, in case
        // the static ctor missed it.
        public override int Execute(params string[] parameters)
        {
            try
            {
                if (_uiContext == null)
                {
                    _uiContext = SynchronizationContext.Current
                                 ?? new WindowsFormsSynchronizationContext();
                    Log("Captured UI SynchronizationContext from Execute().");
                }

                TryStartListener();

                string status = (_listener != null && _listener.IsListening)
                    ? $"listening on http://localhost:{Port}/"
                    : "NOT listening (see log)";

                MessageBox.Show(
                    $"Navisworks MCP Bridge {Version}\r\n\r\n" +
                    $"Status: {status}\r\n" +
                    $"Routes: 35\r\n" +
                    $"Log:    {LogPath}",
                    "Navisworks MCP Bridge",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("Execute() failed: " + ex);
            }
            return 0;
        }

        // ─────────────────────────────────────────────────────────────────────
        //  HTTP listener
        // ─────────────────────────────────────────────────────────────────────

        private static void TryStartListener()
        {
            lock (_startLock)
            {
                if (_listener != null && _listener.IsListening) return;

                try
                {
                    _listener = new HttpListener();
                    _listener.Prefixes.Add($"http://localhost:{Port}/");
                    _listener.Start();

                    _listenerThread = new Thread(ListenLoop)
                    {
                        IsBackground = true,
                        Name = "MCPBridge-HttpListener"
                    };
                    _listenerThread.Start();

                    Log($"Listener started on http://localhost:{Port}/");
                }
                catch (Exception ex)
                {
                    Log("Listener failed to start: " + ex);
                    _listener = null;
                }
            }
        }

        private static void ListenLoop()
        {
            while (_listener != null && _listener.IsListening)
            {
                HttpListenerContext ctx;
                try
                {
                    ctx = _listener.GetContext();
                }
                catch (HttpListenerException) { break; }    // listener stopped
                catch (ObjectDisposedException) { break; }
                catch (Exception ex)
                {
                    Log("GetContext error: " + ex);
                    continue;
                }

                ThreadPool.QueueUserWorkItem(_ => HandleRequest(ctx));
            }
        }

        private static void HandleRequest(HttpListenerContext ctx)
        {
            string path = (ctx.Request.Url?.AbsolutePath ?? "/")
                          .TrimEnd('/')
                          .ToLowerInvariant();
            if (path.Length == 0) path = "/";

            try
            {
                Dictionary<string, object> body = ReadJsonBody(ctx.Request);
                Log($"→ {ctx.Request.HttpMethod} {path}");

                object result = Dispatch(path, body);
                Respond(ctx, 200, result);
            }
            catch (Exception ex)
            {
                Log($"✗ {path}: {ex.GetType().Name}: {ex.Message}");
                try
                {
                    Respond(ctx, 500, new
                    {
                        error = ex.GetType().Name,
                        detail = ex.Message,
                        path
                    });
                }
                catch { /* response already half-written */ }
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        //  Routing
        // ─────────────────────────────────────────────────────────────────────

        private static object Dispatch(string path, Dictionary<string, object> body)
        {
            // Health needs neither UI thread nor active document.
            if (path == "/health")
            {
                return new
                {
                    ok = true,
                    version = Version,
                    port = Port,
                    routes = 35
                };
            }

            // Everything else marshals onto the Navisworks UI thread.
            return RunOnUi(() =>
            {
                switch (path)
                {
                    // Document
                    case "/document/info": return Handlers.DocumentInfo();
                    case "/document/statistics": return Handlers.DocumentStatistics();

                    // Selection
                    case "/selection/get": return Handlers.SelectionGet();
                    case "/selection/select-all": return Handlers.SelectionSelectAll();
                    case "/selection/deselect-all": return Handlers.SelectionDeselectAll();
                    case "/selection/invert": return Handlers.SelectionInvert();
                    case "/selection/from-set": return Handlers.SelectionFromSet(body);
                    case "/selection/by-name": return Handlers.SelectionByName(body);

                    // Sets
                    case "/sets/list": return Handlers.SetsList();
                    case "/sets/selection/create": return Handlers.SelectionSetCreate(body);
                    case "/sets/search/create": return Handlers.SearchSetCreate(body);
                    case "/sets/delete": return Handlers.SetDelete(body);
                    case "/sets/items": return Handlers.SetItems(body);
                    case "/sets/search/run": return Handlers.SearchSetRun(body);

                    // Find / properties
                    case "/elements/find": return Handlers.ElementsFind(body);
                    case "/elements/find-by-name": return Handlers.ElementsFindByName(body);
                    case "/elements/properties": return Handlers.ElementsProperties(body);

                    // Overrides
                    case "/elements/color": return Handlers.ElementsColor(body);
                    case "/elements/transparency": return Handlers.ElementsTransparency(body);
                    case "/elements/hide": return Handlers.ElementsHide();
                    case "/elements/unhide-all": return Handlers.ElementsUnhideAll();
                    case "/elements/isolate": return Handlers.ElementsIsolate();
                    case "/elements/reset-overrides": return Handlers.ElementsResetOverrides();

                    // Clash
                    case "/clash/list": return Handlers.ClashList();
                    case "/clash/create": return Handlers.ClashCreate(body);
                    case "/clash/set-selections": return Handlers.ClashSetSelections(body);
                    case "/clash/run": return Handlers.ClashRun(body);
                    case "/clash/results": return Handlers.ClashResults(body);
                    case "/clash/delete": return Handlers.ClashDelete(body);
                    case "/clash/group/create": return Handlers.ClashGroupCreate(body);
                    case "/clash/group/list": return Handlers.ClashGroupList(body);
                    case "/clash/group/delete": return Handlers.ClashGroupDelete(body);
                    case "/clash/group/add-clashes": return Handlers.ClashGroupAddClashes(body);
                    case "/clash/group/set-status": return Handlers.ClashGroupSetStatus(body);

                    // Viewpoints
                    case "/viewpoints/list": return Handlers.ViewpointsList();
                    case "/viewpoints/save": return Handlers.ViewpointSave(body);
                    case "/viewpoints/goto": return Handlers.ViewpointGoto(body);
                    case "/viewpoints/delete": return Handlers.ViewpointDelete(body);
                    case "/viewpoints/current": return Handlers.ViewpointCurrent();

                    default:
                        throw new InvalidOperationException($"unknown route: {path}");
                }
            });
        }

        // ─────────────────────────────────────────────────────────────────────
        //  UI-thread marshaling
        // ─────────────────────────────────────────────────────────────────────

        private static T RunOnUi<T>(Func<T> func)
        {
            if (_uiContext == null)
            {
                throw new InvalidOperationException(
                    "UI SynchronizationContext not captured. " +
                    "Click the 'MCP Bridge' ribbon button once to initialize.");
            }

            T result = default;
            Exception captured = null;

            _uiContext.Send(_ =>
            {
                try { result = func(); }
                catch (Exception ex) { captured = ex; }
            }, null);

            if (captured != null)
            {
                throw new Exception(captured.Message, captured);
            }
            return result;
        }

        // ─────────────────────────────────────────────────────────────────────
        //  HTTP I/O helpers
        // ─────────────────────────────────────────────────────────────────────

        private static Dictionary<string, object> ReadJsonBody(HttpListenerRequest req)
        {
            if (!req.HasEntityBody) return new Dictionary<string, object>();

            using (var sr = new StreamReader(req.InputStream, req.ContentEncoding ?? Encoding.UTF8))
            {
                string text = sr.ReadToEnd();
                if (string.IsNullOrWhiteSpace(text)) return new Dictionary<string, object>();
                return Json.ParseObject(text);
            }
        }

        private static void Respond(HttpListenerContext ctx, int statusCode, object payload)
        {
            byte[] data = Encoding.UTF8.GetBytes(Json.Serialize(payload));
            ctx.Response.StatusCode = statusCode;
            ctx.Response.ContentType = "application/json; charset=utf-8";
            ctx.Response.ContentLength64 = data.Length;
            ctx.Response.OutputStream.Write(data, 0, data.Length);
            ctx.Response.OutputStream.Close();
        }

        // ─────────────────────────────────────────────────────────────────────
        //  Logging
        // ─────────────────────────────────────────────────────────────────────

        internal static void Log(string message)
        {
            try
            {
                File.AppendAllText(LogPath, $"[{DateTime.Now:O}] {message}\r\n");
            }
            catch { /* never let a logging error kill the plugin */ }
        }
    }

    // =========================================================================
    //  Route handlers
    // =========================================================================
    //  Each method is invoked on the Navisworks UI thread and may freely call
    //  the Navisworks API. Throw on bad input or missing document; the
    //  dispatcher converts the exception into a 500 JSON response.
    // =========================================================================
    internal static partial class Handlers
    {
        // ════════════════════════════════════════════════════════════════════
        //  DOCUMENT
        // ════════════════════════════════════════════════════════════════════

        public static object DocumentInfo()
        {
            var doc = RequireDocument();

            var models = new List<object>();
            foreach (Model m in doc.Models)
            {
                models.Add(new
                {
                    name = m.RootItem?.DisplayName ?? "",
                    fileName = m.FileName ?? "",
                    sourceFile = m.SourceFileName ?? ""
                });
            }

            return new
            {
                title = doc.Title,
                fileName = doc.FileName,
                units = doc.Units.ToString(),
                isClear = doc.IsClear,
                modelCount = doc.Models.Count,
                models
            };
        }

        public static object DocumentStatistics()
        {
            var doc = RequireDocument();

            var counts = new Dictionary<string, int>();
            int total = 0;

            foreach (ModelItem mi in AllItems(doc))
            {
                total++;
                string key = mi.ClassDisplayName ?? "(none)";
                counts.TryGetValue(key, out int c);
                counts[key] = c + 1;
            }

            return new
            {
                total,
                byClass = counts.OrderByDescending(kv => kv.Value)
                                .Select(kv => new { className = kv.Key, count = kv.Value })
                                .ToList()
            };
        }

        // ════════════════════════════════════════════════════════════════════
        //  SELECTION
        // ════════════════════════════════════════════════════════════════════

        public static object SelectionGet()
        {
            var doc = RequireDocument();
            int limit = 100;

            var selected = doc.CurrentSelection.SelectedItems;
            var sample = selected.Take(limit).Select(it => new
            {
                displayName = it.DisplayName,
                className = it.ClassDisplayName
            }).ToList();

            return new
            {
                count = selected.Count,
                sampled = sample.Count,
                items = sample
            };
        }

        public static object SelectionSelectAll()
        {
            var doc = RequireDocument();

            var all = new ModelItemCollection();
            foreach (ModelItem mi in AllItems(doc)) all.Add(mi);

            doc.CurrentSelection.CopyFrom(all);

            return new { selected = all.Count };
        }

        public static object SelectionDeselectAll()
        {
            var doc = RequireDocument();
            doc.CurrentSelection.Clear();
            return new { ok = true };
        }

        public static object SelectionInvert()
        {
            var doc = RequireDocument();

            var current = new HashSet<ModelItem>(doc.CurrentSelection.SelectedItems);
            var inverted = new ModelItemCollection();
            foreach (ModelItem mi in AllItems(doc))
            {
                if (!current.Contains(mi)) inverted.Add(mi);
            }
            doc.CurrentSelection.CopyFrom(inverted);

            return new { selected = inverted.Count };
        }

        public static object SelectionFromSet(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            string name = GetString(body, "name", null)
                ?? throw new ArgumentException("name is required");

            ModelItemCollection items = ResolveNamedSet(doc, name);
            doc.CurrentSelection.CopyFrom(items);

            return new { name, selected = items.Count };
        }

        public static object SelectionByName(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            string pattern = GetString(body, "pattern", null)
                ?? throw new ArgumentException("pattern is required");
            bool exact = GetBool(body, "exact", false);

            var found = new ModelItemCollection();
            foreach (ModelItem mi in AllItems(doc))
            {
                string n = mi.DisplayName ?? "";
                if (exact)
                {
                    if (string.Equals(n, pattern, StringComparison.OrdinalIgnoreCase))
                        found.Add(mi);
                }
                else
                {
                    if (n.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0)
                        found.Add(mi);
                }
            }
            doc.CurrentSelection.CopyFrom(found);

            return new { pattern, selected = found.Count };
        }

        // ════════════════════════════════════════════════════════════════════
        //  SETS
        // ════════════════════════════════════════════════════════════════════

        public static object SetsList()
        {
            var doc = RequireDocument();

            var list = new List<object>();
            void Walk(GroupItem group, string prefix)
            {
                foreach (SavedItem si in group.Children)
                {
                    if (si is SelectionSet set)
                    {
                        // Modern Navisworks unifies selection and search sets under
                        // SelectionSet. Search==null means explicit; non-null means
                        // search-based (the search will be re-evaluated when used).
                        if (!set.HasSearch)
                        {
                            list.Add(new
                            {
                                name = prefix + set.DisplayName,
                                type = "selection",
                                count = set.ExplicitModelItems.Count
                            });
                        }
                        else
                        {
                            list.Add(new
                            {
                                name = prefix + set.DisplayName,
                                type = "search"
                            });
                        }
                    }
                    else if (si is GroupItem g)
                    {
                        Walk(g, prefix + g.DisplayName + "/");
                    }
                }
            }
            Walk(doc.SelectionSets.RootItem, "");

            return new { count = list.Count, sets = list };
        }

        public static object SelectionSetCreate(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            string name = GetString(body, "name", null)
                ?? throw new ArgumentException("name is required");

            var current = doc.CurrentSelection.SelectedItems;
            if (current.Count == 0)
            {
                throw new InvalidOperationException("nothing is currently selected");
            }

            // Construct an explicit selection set and copy items in.
            var set = new SelectionSet { DisplayName = name };
            set.ExplicitModelItems.CopyFrom(current);
            doc.SelectionSets.AddCopy(set);

            return new { name, count = current.Count };
        }

        public static object SearchSetCreate(Dictionary<string, object> body)
        {
            var doc = RequireDocument();

            string name = GetString(body, "name", null)
                ?? throw new ArgumentException("name is required");
            string category = GetString(body, "category", "Item");
            string property = GetString(body, "property", "Name");
            string op = GetString(body, "operator", "contains");
            string value = GetString(body, "value", "");

            var search = new Search();
            search.Selection.SelectAll();
            search.Locations = SearchLocations.DescendantsAndSelf;
            search.SearchConditions.Add(BuildCondition(category, property, op, value));

            // Modern Navisworks: a SelectionSet constructed from a Search IS a search set.
            var searchSet = new SelectionSet(search) { DisplayName = name };
            doc.SelectionSets.AddCopy(searchSet);

            ModelItemCollection matches = search.FindAll(doc, false);

            return new { name, matches = matches.Count };
        }

        public static object SetDelete(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            string name = GetString(body, "name", null)
                ?? throw new ArgumentException("name is required");

            SavedItem found = FindSetByName(doc.SelectionSets.RootItem, name)
                ?? throw new InvalidOperationException($"set '{name}' not found");

            doc.SelectionSets.Remove(found);

            return new { name, deleted = true };
        }

        public static object SetItems(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            string name = GetString(body, "name", null)
                ?? throw new ArgumentException("name is required");

            ModelItemCollection items = ResolveNamedSet(doc, name);

            int limit = GetInt(body, "limit", 100);
            var sample = items.Take(limit).Select(it => new
            {
                displayName = it.DisplayName,
                className = it.ClassDisplayName
            }).ToList();

            return new
            {
                name,
                count = items.Count,
                sampled = sample.Count,
                items = sample
            };
        }

        public static object SearchSetRun(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            string name = GetString(body, "name", null)
                ?? throw new ArgumentException("name is required");

            SavedItem found = FindSetByName(doc.SelectionSets.RootItem, name)
                ?? throw new InvalidOperationException($"set '{name}' not found");

            if (!(found is SelectionSet ss))
            {
                throw new InvalidOperationException($"'{name}' is not a selection set");
            }
            if (!ss.HasSearch)
            {
                throw new InvalidOperationException(
                    $"'{name}' is an explicit selection set, not a search set");
            }

            ModelItemCollection items = ss.Search.FindAll(doc, false);
            return new { name, matches = items.Count };
        }

        // ════════════════════════════════════════════════════════════════════
        //  FIND / PROPERTIES
        // ════════════════════════════════════════════════════════════════════

        public static object ElementsFind(Dictionary<string, object> body)
        {
            var doc = RequireDocument();

            string category = GetString(body, "category", "Item");
            string property = GetString(body, "property", "Name");
            string op = GetString(body, "operator", "contains");
            string value = GetString(body, "value", "");
            int limit = GetInt(body, "limit", 50);
            bool select = GetBool(body, "select", false);

            var search = new Search();
            search.Selection.SelectAll();
            search.Locations = SearchLocations.DescendantsAndSelf;
            search.SearchConditions.Add(BuildCondition(category, property, op, value));

            ModelItemCollection results = search.FindAll(doc, false);

            if (select) doc.CurrentSelection.CopyFrom(results);

            var sample = results.Take(limit).Select(it => new
            {
                displayName = it.DisplayName,
                className = it.ClassDisplayName
            }).ToList();

            return new
            {
                count = results.Count,
                sampled = sample.Count,
                items = sample,
                selected = select
            };
        }

        public static object ElementsFindByName(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            string pattern = GetString(body, "pattern", null)
                ?? throw new ArgumentException("pattern is required");
            int limit = GetInt(body, "limit", 50);
            bool select = GetBool(body, "select", false);

            var found = new ModelItemCollection();
            foreach (ModelItem mi in AllItems(doc))
            {
                string n = mi.DisplayName ?? "";
                if (n.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    found.Add(mi);
                }
            }

            if (select) doc.CurrentSelection.CopyFrom(found);

            var sample = found.Take(limit).Select(it => new
            {
                displayName = it.DisplayName,
                className = it.ClassDisplayName
            }).ToList();

            return new
            {
                pattern,
                count = found.Count,
                sampled = sample.Count,
                items = sample,
                selected = select
            };
        }

        public static object ElementsProperties(Dictionary<string, object> body)
        {
            var doc = RequireDocument();

            // Source: by name (search), or first selected item.
            ModelItem target = null;
            string name = GetString(body, "name", null);
            if (!string.IsNullOrEmpty(name))
            {
                foreach (ModelItem mi in AllItems(doc))
                {
                    if (string.Equals(mi.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                    {
                        target = mi;
                        break;
                    }
                }
                if (target == null)
                {
                    throw new InvalidOperationException($"item '{name}' not found");
                }
            }
            else
            {
                var sel = doc.CurrentSelection.SelectedItems;
                if (sel.Count == 0)
                {
                    throw new InvalidOperationException(
                        "no item selected and no name provided");
                }
                target = sel.First();
            }

            var categories = new List<object>();
            foreach (PropertyCategory cat in target.PropertyCategories)
            {
                var props = new List<object>();
                foreach (DataProperty p in cat.Properties)
                {
                    string v;
                    try { v = p.Value.ToDisplayString(); }
                    catch { v = "(unreadable)"; }

                    props.Add(new { name = p.DisplayName, value = v });
                }
                categories.Add(new { name = cat.DisplayName, properties = props });
            }

            return new
            {
                displayName = target.DisplayName,
                className = target.ClassDisplayName,
                categories
            };
        }

        // ════════════════════════════════════════════════════════════════════
        //  OVERRIDES
        // ════════════════════════════════════════════════════════════════════

        public static object ElementsColor(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            var items = doc.CurrentSelection.SelectedItems;
            if (items.Count == 0)
            {
                throw new InvalidOperationException("nothing is currently selected");
            }

            int r = GetInt(body, "r", 255);
            int g = GetInt(body, "g", 0);
            int b = GetInt(body, "b", 0);

            // Navisworks Color components are 0..1 floats.
            var color = new NwColor(
                Clamp01(r / 255.0),
                Clamp01(g / 255.0),
                Clamp01(b / 255.0));

            doc.Models.OverridePermanentColor(items, color);

            return new { colored = items.Count, r, g, b };
        }

        public static object ElementsTransparency(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            var items = doc.CurrentSelection.SelectedItems;
            if (items.Count == 0)
            {
                throw new InvalidOperationException("nothing is currently selected");
            }

            double t = Clamp01(GetDouble(body, "transparency", 0.5));
            doc.Models.OverridePermanentTransparency(items, t);

            return new { applied = items.Count, transparency = t };
        }

        public static object ElementsHide()
        {
            var doc = RequireDocument();
            var items = doc.CurrentSelection.SelectedItems;
            if (items.Count == 0)
            {
                throw new InvalidOperationException("nothing is currently selected");
            }

            doc.Models.SetHidden(items, true);
            return new { hidden = items.Count };
        }

        public static object ElementsUnhideAll()
        {
            var doc = RequireDocument();

            var all = new ModelItemCollection();
            foreach (ModelItem mi in AllItems(doc)) all.Add(mi);

            doc.Models.SetHidden(all, false);
            return new { unhidden = all.Count };
        }

        public static object ElementsIsolate()
        {
            var doc = RequireDocument();
            var selected = doc.CurrentSelection.SelectedItems;
            if (selected.Count == 0)
            {
                throw new InvalidOperationException("nothing is currently selected");
            }

            var keep = new HashSet<ModelItem>(selected);
            var all = new ModelItemCollection();
            var toHide = new ModelItemCollection();
            foreach (ModelItem mi in AllItems(doc))
            {
                all.Add(mi);
                if (!keep.Contains(mi)) toHide.Add(mi);
            }

            // Show everything first to clear stale hidden state, then hide non-selected.
            doc.Models.SetHidden(all, false);
            doc.Models.SetHidden(toHide, true);

            return new { isolated = selected.Count, hidden = toHide.Count };
        }

        public static object ElementsResetOverrides()
        {
            var doc = RequireDocument();
            doc.Models.ResetAllPermanentMaterials();

            // Also unhide everything for a clean slate.
            var all = new ModelItemCollection();
            foreach (ModelItem mi in AllItems(doc)) all.Add(mi);
            doc.Models.SetHidden(all, false);

            return new { reset = true };
        }

        // ════════════════════════════════════════════════════════════════════
        //  CLASH
        // ════════════════════════════════════════════════════════════════════

        public static object ClashList()
        {
            var doc = RequireDocument();
            var dc = doc.GetClash();

            var tests = new List<object>();
            foreach (ClashTest t in EnumerateClashTests(dc.TestsData.Value.TestsRoot))
            {
                int resultCount = 0;
                CountClashResults(t, ref resultCount);

                tests.Add(new
                {
                    name = t.DisplayName,
                    type = t.TestType.ToString(),
                    tolerance = t.Tolerance,
                    status = t.Status.ToString(),
                    results = resultCount
                });
            }

            return new { count = tests.Count, tests };
        }

        public static object ClashCreate(Dictionary<string, object> body)
        {
            var doc = RequireDocument();

            string name = GetString(body, "name", "Test " + DateTime.Now.ToString("yyyyMMdd-HHmmss"));
            string typeStr = GetString(body, "type", "Hard");
            double tolerance = GetDouble(body, "tolerance", 0.001);
            string setA = GetString(body, "selectionA", null);
            string setB = GetString(body, "selectionB", null);

            DocumentClash dc = doc.GetClash();

            var test = new ClashTest
            {
                DisplayName = name,
                TestType = ParseClashType(typeStr),
                Tolerance = tolerance
            };

            int itemsA = 0, itemsB = 0;
            if (!string.IsNullOrEmpty(setA))
            {
                var items = ResolveNamedSet(doc, setA);
                // Selection has CopyFrom(ModelItemCollection) — confirmed via reflection.
                test.SelectionA.Selection.CopyFrom(items);
                itemsA = items.Count;
            }
            if (!string.IsNullOrEmpty(setB))
            {
                var items = ResolveNamedSet(doc, setB);
                test.SelectionB.Selection.CopyFrom(items);
                itemsB = items.Count;
            }

            // TestsAddCopy needs the root ClashTestFolder (a GroupItem).
            dc.TestsData.TestsAddCopy(dc.TestsData.Value.TestsRoot, test);

            return new
            {
                name = test.DisplayName,
                type = test.TestType.ToString(),
                tolerance = test.Tolerance,
                selectionA = setA,
                selectionB = setB,
                itemsA,
                itemsB
            };
        }

        // Update the A/B selections on an EXISTING clash test by name. Lets you
        // attach sets to tests that were created without them, without having
        // to delete and recreate the test (which loses any results history).
        public static object ClashSetSelections(Dictionary<string, object> body)
        {
            var doc = RequireDocument();

            string name = GetString(body, "name", null)
                ?? throw new ArgumentException("name is required");
            string setA = GetString(body, "selectionA", null);
            string setB = GetString(body, "selectionB", null);

            if (setA == null && setB == null)
            {
                throw new ArgumentException("at least one of selectionA / selectionB is required");
            }

            DocumentClash dc = doc.GetClash();

            // Resolve sets BEFORE the foreach so we don't risk doing API calls
            // that could invalidate the iteration.
            ModelItemCollection itemsA = setA != null ? ResolveNamedSet(doc, setA) : null;
            ModelItemCollection itemsB = setB != null ? ResolveNamedSet(doc, setB) : null;

            // Capture a stable reference; resolve fresh before mutating.
            SavedItemReference targetRef = null;
            foreach (SavedItem si in dc.TestsData.Value.TestsRoot.Children)
            {
                if (si is ClashTest t &&
                    string.Equals(t.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                {
                    targetRef = dc.TestsData.CreateReference(t);
                    break;
                }
            }
            if (targetRef == null)
            {
                throw new InvalidOperationException($"clash test '{name}' not found");
            }

            var fresh = (ClashTest)dc.TestsData.ResolveReference(targetRef);
            if (itemsA != null) fresh.SelectionA.Selection.CopyFrom(itemsA);
            if (itemsB != null) fresh.SelectionB.Selection.CopyFrom(itemsB);

            return new
            {
                name,
                selectionA = setA,
                selectionB = setB,
                itemsA = itemsA?.Count ?? 0,
                itemsB = itemsB?.Count ?? 0
            };
        }

        public static object ClashRun(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            string name = GetString(body, "name", null);

            DocumentClash dc = doc.GetClash();
            var ran = new List<string>();

            if (string.IsNullOrEmpty(name))
            {
                // Use the dedicated API method — avoids the WeakRef disposal
                // problem we'd hit by iterating + calling TestsRunTest per test.
                dc.TestsData.TestsRunAllTests();

                // Re-iterate to gather names for the response.
                foreach (SavedItem si in dc.TestsData.Value.TestsRoot.Children)
                {
                    if (si is ClashTest t) ran.Add(t.DisplayName);
                }
            }
            else
            {
                // Capture a stable SavedItemReference during iteration, then re-resolve
                // for a fresh handle just before calling TestsRunTest. ClashTest objects
                // obtained directly from Children are weak NativeHandles that go stale
                // after the iteration completes (or sometimes immediately).
                SavedItemReference targetRef = null;
                foreach (SavedItem si in dc.TestsData.Value.TestsRoot.Children)
                {
                    if (si is ClashTest t &&
                        string.Equals(t.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                    {
                        targetRef = dc.TestsData.CreateReference(t);
                        break;
                    }
                }
                if (targetRef == null)
                {
                    throw new InvalidOperationException($"clash test '{name}' not found");
                }

                // Resolve fresh and run immediately. Don't access target.DisplayName
                // afterward — TestsRunTest may invalidate its argument.
                var fresh = (ClashTest)dc.TestsData.ResolveReference(targetRef);
                dc.TestsData.TestsRunTest(fresh);
                ran.Add(name);
            }

            return new { ran };
        }

        public static object ClashResults(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            string name = GetString(body, "name", null)
                ?? throw new ArgumentException("name is required");
            int limit = GetInt(body, "limit", 100);

            DocumentClash dc = doc.GetClash();

            SavedItemReference targetRef = null;
            foreach (SavedItem si in dc.TestsData.Value.TestsRoot.Children)
            {
                if (si is ClashTest t &&
                    string.Equals(t.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                {
                    targetRef = dc.TestsData.CreateReference(t);
                    break;
                }
            }
            if (targetRef == null)
            {
                throw new InvalidOperationException($"clash test '{name}' not found");
            }

            var fresh = (ClashTest)dc.TestsData.ResolveReference(targetRef);
            string status = fresh.Status.ToString();
            var results = new List<object>();
            CollectClashResults(fresh, results);

            return new
            {
                name,
                status,
                count = results.Count,
                sampled = Math.Min(results.Count, limit),
                results = results.Take(limit).ToList()
            };
        }

        public static object ClashDelete(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            string name = GetString(body, "name", null)
                ?? throw new ArgumentException("name is required");

            DocumentClash dc = doc.GetClash();
            ClashTestFolder root = dc.TestsData.Value.TestsRoot;

            // Find the index, then remove. We deliberately don't recurse into
            // sub-groups here — top-level only — because sub-group deletion
            // needs the actual parent group ref and gets into the same WeakRef
            // territory we're trying to avoid.
            int index = 0;
            foreach (SavedItem si in root.Children)
            {
                if (si is ClashTest t &&
                    string.Equals(t.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                {
                    dc.TestsData.TestsRemoveAt(root, index);
                    return new { name, deleted = true };
                }
                index++;
            }
            throw new InvalidOperationException($"clash test '{name}' not found");
        }

        // ════════════════════════════════════════════════════════════════════
        //  VIEWPOINTS
        // ════════════════════════════════════════════════════════════════════

        public static object ViewpointsList()
        {
            var doc = RequireDocument();

            var list = new List<object>();
            void Walk(GroupItem group, string prefix)
            {
                foreach (SavedItem si in group.Children)
                {
                    if (si is SavedViewpoint sv)
                    {
                        list.Add(new { name = prefix + sv.DisplayName, type = "viewpoint" });
                    }
                    else if (si is GroupItem g)
                    {
                        Walk(g, prefix + g.DisplayName + "/");
                    }
                }
            }
            Walk(doc.SavedViewpoints.RootItem, "");

            return new { count = list.Count, viewpoints = list };
        }

        public static object ViewpointSave(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            string name = GetString(body, "name", null)
                ?? throw new ArgumentException("name is required");

            var sv = new SavedViewpoint(doc.CurrentViewpoint.CreateCopy())
            {
                DisplayName = name
            };
            doc.SavedViewpoints.AddCopy(sv);

            return new { name };
        }

        public static object ViewpointGoto(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            string name = GetString(body, "name", null)
                ?? throw new ArgumentException("name is required");

            SavedViewpoint sv = FindSavedViewpoint(doc.SavedViewpoints.RootItem, name)
                ?? throw new InvalidOperationException($"saved viewpoint '{name}' not found");

            doc.CurrentViewpoint.CopyFrom(sv.Viewpoint);

            return new { name = sv.DisplayName };
        }

        public static object ViewpointDelete(Dictionary<string, object> body)
        {
            var doc = RequireDocument();
            string name = GetString(body, "name", null)
                ?? throw new ArgumentException("name is required");

            SavedViewpoint sv = FindSavedViewpoint(doc.SavedViewpoints.RootItem, name)
                ?? throw new InvalidOperationException($"saved viewpoint '{name}' not found");

            doc.SavedViewpoints.Remove(sv);
            return new { name, deleted = true };
        }

        public static object ViewpointCurrent()
        {
            var doc = RequireDocument();
            Viewpoint vp = doc.CurrentViewpoint.CreateCopy();

            return new
            {
                position = new { x = vp.Position.X, y = vp.Position.Y, z = vp.Position.Z },
                projection = vp.Projection.ToString(),
                focalDistance = vp.FocalDistance,
                hasLighting = vp.HasLighting
            };
        }

        // ════════════════════════════════════════════════════════════════════
        //  Internal helpers
        // ════════════════════════════════════════════════════════════════════

        private static Document RequireDocument()
        {
            var doc = NavisApp.ActiveDocument;
            if (doc == null || doc.IsClear)
            {
                throw new InvalidOperationException("no active document");
            }
            return doc;
        }

        // Walks every model item in every model. Use sparingly — large
        // federated models can have hundreds of thousands of items.
        private static IEnumerable<ModelItem> AllItems(Document doc)
        {
            foreach (Model m in doc.Models)
            {
                if (m.RootItem == null) continue;
                foreach (ModelItem mi in m.RootItem.DescendantsAndSelf)
                {
                    yield return mi;
                }
            }
        }

        // Resolve a SelectionSet (explicit or search-based) by name to its current items.
        private static ModelItemCollection ResolveNamedSet(Document doc, string name)
        {
            SavedItem hit = FindSetByName(doc.SelectionSets.RootItem, name);
            if (hit == null)
            {
                throw new InvalidOperationException($"set '{name}' not found");
            }
            if (hit is SelectionSet set)
            {
                if (!set.HasSearch)
                {
                    // Explicit set: items are stored directly.
                    return new ModelItemCollection(set.ExplicitModelItems);
                }
                // Search set: re-evaluate the search.
                return set.Search.FindAll(doc, false);
            }
            throw new InvalidOperationException($"'{name}' is not a selection or search set");
        }

        private static SavedItem FindSetByName(GroupItem root, string name)
        {
            if (root == null) return null;
            foreach (SavedItem si in root.Children)
            {
                if (string.Equals(si.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                {
                    return si;
                }
                if (si is GroupItem g)
                {
                    SavedItem hit = FindSetByName(g, name);
                    if (hit != null) return hit;
                }
            }
            return null;
        }

        private static SavedViewpoint FindSavedViewpoint(GroupItem root, string name)
        {
            if (root == null) return null;
            foreach (SavedItem item in root.Children)
            {
                if (item is SavedViewpoint sv &&
                    string.Equals(sv.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                {
                    return sv;
                }
                if (item is GroupItem g)
                {
                    SavedViewpoint hit = FindSavedViewpoint(g, name);
                    if (hit != null) return hit;
                }
            }
            return null;
        }

        private static SearchCondition BuildCondition(
            string category, string property, string op, string value)
        {
            // Build "<category>/<property> <op> <value>" condition using the
            // display names that the user sees in the Properties panel.
            SearchCondition cond = SearchCondition.HasPropertyByDisplayName(category, property);

            // 2027 SDK: SearchCondition.WildcardValue is gone. EqualValue will
            // interpret '*' as a wildcard inside the comparison string for
            // string-typed properties, which covers contains/startsWith/endsWith.
            switch ((op ?? "").Trim().ToLowerInvariant())
            {
                case "equals":
                case "=":
                    cond = cond.EqualValue(VariantData.FromDisplayString(value));
                    break;
                case "contains":
                case "~":
                    cond = cond.EqualValue(VariantData.FromDisplayString("*" + value + "*"));
                    break;
                case "startswith":
                    cond = cond.EqualValue(VariantData.FromDisplayString(value + "*"));
                    break;
                case "endswith":
                    cond = cond.EqualValue(VariantData.FromDisplayString("*" + value));
                    break;
                default:
                    throw new ArgumentException(
                        $"unknown operator '{op}'. Use equals, contains, startsWith, endsWith.");
            }
            return cond;
        }

        // ── Clash helpers ───────────────────────────────────────────────────

        // Walk every ClashTest under a group recursively (groups can nest
        // in modern Navisworks Clash Detective).
        private static IEnumerable<ClashTest> EnumerateClashTests(GroupItem root)
        {
            if (root == null) yield break;
            foreach (SavedItem si in root.Children)
            {
                if (si is ClashTest t) yield return t;
                else if (si is GroupItem g)
                {
                    foreach (ClashTest nested in EnumerateClashTests(g))
                        yield return nested;
                }
            }
        }

        private static ClashTest FindClashTest(DocumentClash dc, string name)
        {
            foreach (ClashTest t in EnumerateClashTests(dc.TestsData.Value.TestsRoot))
            {
                if (string.Equals(t.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                    return t;
            }
            return null;
        }

        // Find a test plus its immediate parent + index within that parent.
        // Needed for TestsRemoveAt(GroupItem parent, int index) in 2027.
        private static bool TryFindClashTestLocation(
            GroupItem parent, string name,
            out GroupItem foundParent, out int foundIndex)
        {
            int i = 0;
            foreach (SavedItem si in parent.Children)
            {
                if (si is ClashTest t &&
                    string.Equals(t.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                {
                    foundParent = parent;
                    foundIndex = i;
                    return true;
                }
                if (si is GroupItem g &&
                    TryFindClashTestLocation(g, name, out foundParent, out foundIndex))
                {
                    return true;
                }
                i++;
            }
            foundParent = null;
            foundIndex = -1;
            return false;
        }

        private static void CountClashResults(SavedItem item, ref int count)
        {
            if (item is ClashResult) count++;
            else if (item is GroupItem g)
            {
                foreach (SavedItem c in g.Children) CountClashResults(c, ref count);
            }
        }

        private static void CollectClashResults(SavedItem item, List<object> output)
        {
            if (item is ClashResult r)
            {
                output.Add(new
                {
                    name = r.DisplayName,
                    status = r.Status.ToString(),
                    distance = r.Distance,
                    item1 = r.Item1?.DisplayName,
                    item2 = r.Item2?.DisplayName
                });
            }
            else if (item is GroupItem g)
            {
                foreach (SavedItem c in g.Children) CollectClashResults(c, output);
            }
        }

        private static ClashTestType ParseClashType(string s)
        {
            if (string.IsNullOrEmpty(s)) return ClashTestType.Hard;
            switch (s.Trim().ToLowerInvariant())
            {
                case "hard": return ClashTestType.Hard;
                case "clearance": return ClashTestType.Clearance;
                case "duplicate": return ClashTestType.Duplicate;
                default:
                    throw new ArgumentException(
                        $"unknown clash type '{s}'. Use Hard, Clearance, or Duplicate.");
            }
        }

        // ── JSON convenience accessors ──────────────────────────────────────
        //  Body comes from the hand-rolled JSON parser as Dictionary<string, object>
        //  with values being string, double, bool, null, List<object>, or nested dict.

        private static string GetString(Dictionary<string, object> body, string key, string @default)
        {
            if (body == null || !body.TryGetValue(key, out object v) || v == null) return @default;
            return v.ToString();
        }

        private static int GetInt(Dictionary<string, object> body, string key, int @default)
        {
            if (body == null || !body.TryGetValue(key, out object v) || v == null) return @default;
            if (v is double d) return (int)d;
            if (v is long l) return (int)l;
            if (v is int i) return i;
            if (int.TryParse(v.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int j)) return j;
            return @default;
        }

        private static double GetDouble(Dictionary<string, object> body, string key, double @default)
        {
            if (body == null || !body.TryGetValue(key, out object v) || v == null) return @default;
            if (v is double d) return d;
            if (v is long l) return (double)l;
            if (v is int i) return (double)i;
            if (double.TryParse(v.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out double e)) return e;
            return @default;
        }

        private static bool GetBool(Dictionary<string, object> body, string key, bool @default)
        {
            if (body == null || !body.TryGetValue(key, out object v) || v == null) return @default;
            if (v is bool b) return b;
            if (bool.TryParse(v.ToString(), out bool r)) return r;
            return @default;
        }

        private static double Clamp01(double x)
        {
            if (x < 0) return 0;
            if (x > 1) return 1;
            return x;
        }
    }

    // =========================================================================
    //  Json — hand-rolled serialization / parsing
    // =========================================================================
    //  We deliberately avoid System.Text.Json (and Newtonsoft) so the plugin
    //  has zero NuGet dependencies. Anything Navisworks happens to load in-
    //  process therefore can't conflict with us. The parser handles top-level
    //  objects with string/number/bool/null/array/object values — enough for
    //  the request bodies our routes accept. The serializer reflects over
    //  anonymous types (which is what every handler returns) plus handles
    //  primitives, dictionaries, and any IEnumerable.
    // =========================================================================
    internal static class Json
    {
        // ── Serialize ───────────────────────────────────────────────────────

        public static string Serialize(object value)
        {
            var sb = new StringBuilder();
            WriteValue(sb, value);
            return sb.ToString();
        }

        private static void WriteValue(StringBuilder sb, object v)
        {
            if (v == null) { sb.Append("null"); return; }

            switch (v)
            {
                case string s: WriteString(sb, s); return;
                case bool b: sb.Append(b ? "true" : "false"); return;
                case int i: sb.Append(i.ToString(CultureInfo.InvariantCulture)); return;
                case long l: sb.Append(l.ToString(CultureInfo.InvariantCulture)); return;
                case double d: sb.Append(FormatDouble(d)); return;
                case float f: sb.Append(FormatDouble(f)); return;
                case decimal m: sb.Append(m.ToString(CultureInfo.InvariantCulture)); return;
                case Enum e: WriteString(sb, e.ToString()); return;
            }

            // Generic dictionary (e.g. Dictionary<string, object>)
            if (v is System.Collections.IDictionary dict)
            {
                WriteDict(sb, dict);
                return;
            }
            // Any IEnumerable (List<>, arrays, IEnumerable<>, etc.)
            if (v is System.Collections.IEnumerable seq)
            {
                WriteSeq(sb, seq);
                return;
            }
            // Anonymous types and POCOs: reflect over public properties.
            WriteObject(sb, v);
        }

        private static string FormatDouble(double d)
        {
            // R round-trips, but emits "Infinity"/"NaN" which aren't valid JSON.
            // Replace those with null-ish values to keep responses parseable.
            if (double.IsNaN(d) || double.IsInfinity(d)) return "null";
            string s = d.ToString("R", CultureInfo.InvariantCulture);
            // ensure decimal point so 1 vs 1.0 distinction is preserved... actually
            // JSON doesn't care, leave as-is.
            return s;
        }

        private static void WriteString(StringBuilder sb, string s)
        {
            sb.Append('"');
            foreach (char c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < 0x20) sb.AppendFormat(CultureInfo.InvariantCulture, "\\u{0:x4}", (int)c);
                        else sb.Append(c);
                        break;
                }
            }
            sb.Append('"');
        }

        private static void WriteDict(StringBuilder sb, System.Collections.IDictionary dict)
        {
            sb.Append('{');
            bool first = true;
            foreach (System.Collections.DictionaryEntry kv in dict)
            {
                if (kv.Value == null) continue;       // mimic ignore-null behavior
                if (!first) sb.Append(',');
                first = false;
                WriteString(sb, kv.Key?.ToString() ?? "");
                sb.Append(':');
                WriteValue(sb, kv.Value);
            }
            sb.Append('}');
        }

        private static void WriteSeq(StringBuilder sb, System.Collections.IEnumerable seq)
        {
            sb.Append('[');
            bool first = true;
            foreach (object item in seq)
            {
                if (!first) sb.Append(',');
                first = false;
                WriteValue(sb, item);
            }
            sb.Append(']');
        }

        private static void WriteObject(StringBuilder sb, object obj)
        {
            sb.Append('{');
            bool first = true;
            var props = obj.GetType().GetProperties(
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            foreach (var p in props)
            {
                object val;
                try { val = p.GetValue(obj); }
                catch { continue; }
                if (val == null) continue;            // mimic ignore-null behavior

                if (!first) sb.Append(',');
                first = false;
                WriteString(sb, p.Name);
                sb.Append(':');
                WriteValue(sb, val);
            }
            sb.Append('}');
        }

        // ── Parse ───────────────────────────────────────────────────────────

        public static Dictionary<string, object> ParseObject(string text)
        {
            int i = 0;
            object root = ReadValue(text, ref i);
            return root as Dictionary<string, object> ?? new Dictionary<string, object>();
        }

        private static object ReadValue(string s, ref int i)
        {
            SkipWs(s, ref i);
            if (i >= s.Length) return null;

            char c = s[i];
            if (c == '{') return ReadObj(s, ref i);
            if (c == '[') return ReadArr(s, ref i);
            if (c == '"') return ReadStr(s, ref i);
            if (c == 't' || c == 'f') return ReadBool(s, ref i);
            if (c == 'n') { i = Math.Min(i + 4, s.Length); return null; }
            return ReadNum(s, ref i);
        }

        private static Dictionary<string, object> ReadObj(string s, ref int i)
        {
            var d = new Dictionary<string, object>();
            i++; // skip '{'
            SkipWs(s, ref i);
            if (i < s.Length && s[i] == '}') { i++; return d; }

            while (i < s.Length)
            {
                SkipWs(s, ref i);
                if (i >= s.Length || s[i] != '"') break;
                string key = ReadStr(s, ref i);
                SkipWs(s, ref i);
                if (i < s.Length && s[i] == ':') i++;
                object val = ReadValue(s, ref i);
                d[key] = val;
                SkipWs(s, ref i);
                if (i < s.Length && s[i] == ',') { i++; continue; }
                if (i < s.Length && s[i] == '}') { i++; break; }
                break;
            }
            return d;
        }

        private static List<object> ReadArr(string s, ref int i)
        {
            var list = new List<object>();
            i++; // skip '['
            SkipWs(s, ref i);
            if (i < s.Length && s[i] == ']') { i++; return list; }

            while (i < s.Length)
            {
                list.Add(ReadValue(s, ref i));
                SkipWs(s, ref i);
                if (i < s.Length && s[i] == ',') { i++; continue; }
                if (i < s.Length && s[i] == ']') { i++; break; }
                break;
            }
            return list;
        }

        private static string ReadStr(string s, ref int i)
        {
            var sb = new StringBuilder();
            i++; // opening "
            while (i < s.Length && s[i] != '"')
            {
                if (s[i] == '\\' && i + 1 < s.Length)
                {
                    char esc = s[i + 1];
                    switch (esc)
                    {
                        case '"': sb.Append('"'); i += 2; break;
                        case '\\': sb.Append('\\'); i += 2; break;
                        case '/': sb.Append('/'); i += 2; break;
                        case 'b': sb.Append('\b'); i += 2; break;
                        case 'f': sb.Append('\f'); i += 2; break;
                        case 'n': sb.Append('\n'); i += 2; break;
                        case 'r': sb.Append('\r'); i += 2; break;
                        case 't': sb.Append('\t'); i += 2; break;
                        case 'u':
                            if (i + 5 < s.Length)
                            {
                                int code = Convert.ToInt32(s.Substring(i + 2, 4), 16);
                                sb.Append((char)code);
                                i += 6;
                            }
                            else { i += 2; }
                            break;
                        default: sb.Append(esc); i += 2; break;
                    }
                }
                else
                {
                    sb.Append(s[i++]);
                }
            }
            if (i < s.Length) i++; // closing "
            return sb.ToString();
        }

        private static bool ReadBool(string s, ref int i)
        {
            if (s[i] == 't') { i = Math.Min(i + 4, s.Length); return true; }
            i = Math.Min(i + 5, s.Length); return false;
        }

        private static object ReadNum(string s, ref int i)
        {
            int start = i;
            while (i < s.Length && (char.IsDigit(s[i]) || "+-.eE".IndexOf(s[i]) >= 0)) i++;
            string token = s.Substring(start, i - start);
            if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out double d))
                return d;
            return 0.0;
        }

        private static void SkipWs(string s, ref int i)
        {
            while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
        }
    }
}
