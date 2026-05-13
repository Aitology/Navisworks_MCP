// =============================================================================
//  ClashGroupHandlers.cs — drop-in file, compiles on its own.
// =============================================================================
//  Extends the existing Handlers class with three new methods:
//      Handlers.ClashGroupCreate
//      Handlers.ClashGroupList
//      Handlers.ClashGroupDelete
//
//  Requirements before this file will compile:
//    1.  In your main .cs, change
//             internal static class Handlers
//        to
//             internal static partial class Handlers
//        (just add the word `partial`. Nothing else changes.)
//
//    2.  Add three case lines to the Dispatch() switch — see bottom of this file.
// =============================================================================

    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using Autodesk.Navisworks.Api;
    using Autodesk.Navisworks.Api.Clash;


namespace NavisworksMcpAddin
{
    internal static partial class Handlers
    {
        // ════════════════════════════════════════════════════════════════════
        //  CLASH GROUPS
        // ════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Creates an (empty) clash group inside an existing clash test.
        /// Body: { "test_name": "...", "group_name": "..." }
        /// </summary>
        public static object ClashGroupCreate(Dictionary<string, object> body)
        {
            var doc = RequireDocument();

            string testName = GetString(body, "test_name", null)
                ?? throw new ArgumentException("test_name is required");
            string groupName = GetString(body, "group_name", null)
                ?? throw new ArgumentException("group_name is required");

            DocumentClash dc = doc.GetClash();

            // Capture a stable reference to the parent test (WeakRef-safe).
            SavedItemReference testRef = null;
            foreach (SavedItem si in dc.TestsData.Value.TestsRoot.Children)
            {
                if (si is ClashTest t &&
                    string.Equals(t.DisplayName, testName, StringComparison.OrdinalIgnoreCase))
                {
                    testRef = dc.TestsData.CreateReference(t);
                    break;
                }
            }
            if (testRef == null)
            {
                throw new InvalidOperationException($"clash test '{testName}' not found");
            }

            // Refuse to create a duplicate group name inside the same test.
            var freshTest = (ClashTest)dc.TestsData.ResolveReference(testRef);
            foreach (SavedItem child in freshTest.Children)
            {
                if (child is ClashResultGroup existing &&
                    string.Equals(existing.DisplayName, groupName, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        $"a clash group named '{groupName}' already exists in test '{testName}'");
                }
            }

            // TestsAddCopy(parent, savedItem) accepts any GroupItem as parent,
            // including a ClashTest, and clones the item in.
            var newGroup = new ClashResultGroup { DisplayName = groupName };
            dc.TestsData.TestsAddCopy(freshTest, newGroup);

            return new
            {
                test_name  = testName,
                group_name = groupName,
                created    = true
            };
        }

        /// <summary>
        /// Lists all clash groups inside the named test, including the names of
        /// the clash results contained in each group.
        /// Body: { "test_name": "..." }
        /// </summary>
        public static object ClashGroupList(Dictionary<string, object> body)
        {
            var doc = RequireDocument();

            string testName = GetString(body, "test_name", null)
                ?? throw new ArgumentException("test_name is required");

            DocumentClash dc = doc.GetClash();

            ClashTest test = null;
            foreach (SavedItem si in dc.TestsData.Value.TestsRoot.Children)
            {
                if (si is ClashTest t &&
                    string.Equals(t.DisplayName, testName, StringComparison.OrdinalIgnoreCase))
                {
                    test = t;
                    break;
                }
            }
            if (test == null)
            {
                throw new InvalidOperationException($"clash test '{testName}' not found");
            }

            var groups = new List<object>();
            foreach (SavedItem child in test.Children)
            {
                if (child is ClashResultGroup g)
                {
                    var clashNames = new List<string>();
                    foreach (SavedItem gc in g.Children)
                    {
                        if (gc is ClashResult cr) clashNames.Add(cr.DisplayName);
                    }

                    groups.Add(new
                    {
                        name        = g.DisplayName,
                        status      = g.Status.ToString(),
                        clash_count = clashNames.Count,
                        clashes     = clashNames
                    });
                }
            }

            return new
            {
                test_name = testName,
                count     = groups.Count,
                groups
            };
        }

        /// <summary>
        /// Deletes a clash group (and its child results) from a test.
        /// Body: { "test_name": "...", "group_name": "..." }
        /// </summary>
        public static object ClashGroupDelete(Dictionary<string, object> body)
        {
            var doc = RequireDocument();

            string testName = GetString(body, "test_name", null)
                ?? throw new ArgumentException("test_name is required");
            string groupName = GetString(body, "group_name", null)
                ?? throw new ArgumentException("group_name is required");

            DocumentClash dc = doc.GetClash();

            SavedItemReference testRef = null;
            foreach (SavedItem si in dc.TestsData.Value.TestsRoot.Children)
            {
                if (si is ClashTest t &&
                    string.Equals(t.DisplayName, testName, StringComparison.OrdinalIgnoreCase))
                {
                    testRef = dc.TestsData.CreateReference(t);
                    break;
                }
            }
            if (testRef == null)
            {
                throw new InvalidOperationException($"clash test '{testName}' not found");
            }

            var freshTest = (ClashTest)dc.TestsData.ResolveReference(testRef);

            int idx = 0;
            int foundIndex = -1;
            foreach (SavedItem child in freshTest.Children)
            {
                if (child is ClashResultGroup g &&
                    string.Equals(g.DisplayName, groupName, StringComparison.OrdinalIgnoreCase))
                {
                    foundIndex = idx;
                    break;
                }
                idx++;
            }
            if (foundIndex == -1)
            {
                throw new InvalidOperationException(
                    $"clash group '{groupName}' not found inside test '{testName}'");
            }

            // TestsRemoveAt(parent, index) — same call shape ClashDelete uses,
            // but with the ClashTest itself as the parent.
            dc.TestsData.TestsRemoveAt(freshTest, foundIndex);

            return new
            {
                test_name  = testName,
                group_name = groupName,
                deleted    = true
            };
        }


        /// <summary>
        /// Moves existing (ungrouped) clash results into an existing clash group.
        /// Body: {
        ///   "test_name":   "...",
        ///   "group_name":  "...",
        ///   "clash_names": ["Clash1","Clash2",...]
        /// }
        /// Navisworks has no direct "move" API for clashes, so we do
        /// add-copy-into-group followed by remove-from-test, re-resolving
        /// handles before every mutation to avoid the WeakRef disposal issue
        /// the rest of the clash code already documents.
        /// </summary>
        public static object ClashGroupAddClashes(Dictionary<string, object> body)
        {
            var doc = RequireDocument();

            string testName = GetString(body, "test_name", null)
                ?? throw new ArgumentException("test_name is required");
            string groupName = GetString(body, "group_name", null)
                ?? throw new ArgumentException("group_name is required");

            if (!body.TryGetValue("clash_names", out object rawNames) ||
                !(rawNames is System.Collections.IEnumerable rawSeq) ||
                rawNames is string)
            {
                throw new ArgumentException("clash_names is required (array of strings)");
            }

            var wanted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var o in rawSeq)
            {
                if (o == null) continue;
                wanted.Add(o.ToString());
            }
            if (wanted.Count == 0)
            {
                throw new ArgumentException("clash_names cannot be empty");
            }

            DocumentClash dc = doc.GetClash();

            // ── Locate parent test (stable reference) ───────────────────────
            SavedItemReference testRef = null;
            foreach (SavedItem si in dc.TestsData.Value.TestsRoot.Children)
            {
                if (si is ClashTest t &&
                    string.Equals(t.DisplayName, testName, StringComparison.OrdinalIgnoreCase))
                {
                    testRef = dc.TestsData.CreateReference(t);
                    break;
                }
            }
            if (testRef == null)
            {
                throw new InvalidOperationException($"clash test '{testName}' not found");
            }

            // ── Locate target group AND collect ordered list of clash names
            //    that actually exist as ungrouped children of this test ─────
            var freshTest = (ClashTest)dc.TestsData.ResolveReference(testRef);
            SavedItemReference groupRef = null;
            var presentNames = new List<string>();  // in test-tree order

            foreach (SavedItem child in freshTest.Children)
            {
                if (child is ClashResultGroup g &&
                    string.Equals(g.DisplayName, groupName, StringComparison.OrdinalIgnoreCase))
                {
                    groupRef = dc.TestsData.CreateReference(g);
                }
                else if (child is ClashResult cr && wanted.Contains(cr.DisplayName))
                {
                    presentNames.Add(cr.DisplayName);
                }
            }

            if (groupRef == null)
            {
                throw new InvalidOperationException(
                    $"clash group '{groupName}' not found in test '{testName}'. " +
                    "Create it first with create_clash_group.");
            }
            if (presentNames.Count == 0)
            {
                throw new InvalidOperationException(
                    $"none of the requested clash names were found as direct " +
                    $"(ungrouped) children of '{testName}'.");
            }

            // ── Per-clash: copy-out, remove-original, add-copy-to-group ────
            //
            //  Key insight: a ClashResult's GUID is its identity. The Clash
            //  test tree refuses to contain two items with the same GUID.
            //
            //  CreateCopy() yields an *orphaned* SavedItem with a new GUID
            //  but identical Item1/Item2/distance/etc. We make this copy
            //  BEFORE removing the original so we still have a live handle
            //  to read from, then add the copy AFTER the remove so the tree
            //  never holds two of the same GUID at once.
            int moved = 0;
            foreach (string name in presentNames)
            {
                // Locate the original (by name) — re-scan each iteration
                // because indices shift after every removal.
                freshTest = (ClashTest)dc.TestsData.ResolveReference(testRef);
                int targetIdx = -1;
                SavedItemReference originalRef = null;

                int j = 0;
                foreach (SavedItem child in freshTest.Children)
                {
                    if (child is ClashResult cr &&
                        string.Equals(cr.DisplayName, name, StringComparison.OrdinalIgnoreCase))
                    {
                        targetIdx = j;
                        originalRef = dc.TestsData.CreateReference(cr);
                        break;
                    }
                    j++;
                }
                if (targetIdx < 0 || originalRef == null) continue; // skip if vanished

                // Step A: build an orphan copy with a fresh GUID
                var originalFresh = (ClashResult)dc.TestsData.ResolveReference(originalRef);
                var orphanCopy = originalFresh.CreateCopy();   // new identity

                // Step B: remove the original (after CreateCopy — order matters)
                freshTest = (ClashTest)dc.TestsData.ResolveReference(testRef);
                dc.TestsData.TestsRemoveAt(freshTest, targetIdx);

                // Step C: add the orphan into the target group
                var freshGroup = (ClashResultGroup)dc.TestsData.ResolveReference(groupRef);
                dc.TestsData.TestsAddCopy(freshGroup, orphanCopy);

                moved++;
            }

            // ── Report ──────────────────────────────────────────────────────
            var notFound = new List<string>();
            foreach (string req in wanted)
            {
                bool got = false;
                foreach (string p in presentNames)
                {
                    if (string.Equals(p, req, StringComparison.OrdinalIgnoreCase))
                    {
                        got = true; break;
                    }
                }
                if (!got) notFound.Add(req);
            }

            return new
            {
                test_name = testName,
                group_name = groupName,
                moved = moved,
                moved_clashes = presentNames,
                not_found = notFound
            };
        }

        /// <summary>
        /// Sets the status of a clash group. Tries the group directly first;
        /// if Navisworks rejects that, falls back to setting status on every
        /// child clash result (which is how the UI does it anyway).
        /// Body: {
        ///   "test_name":  "...",
        ///   "group_name": "...",
        ///   "status":     "New" | "Active" | "Reviewed" | "Approved" | "Resolved"
        /// }
        /// </summary>
        public static object ClashGroupSetStatus(Dictionary<string, object> body)
        {
            var doc = RequireDocument();

            string testName = GetString(body, "test_name", null)
                ?? throw new ArgumentException("test_name is required");
            string groupName = GetString(body, "group_name", null)
                ?? throw new ArgumentException("group_name is required");
            string statusStr = GetString(body, "status", null)
                ?? throw new ArgumentException("status is required");

            ClashResultStatus parsed;
            switch (statusStr.Trim().ToLowerInvariant())
            {
                case "new": parsed = ClashResultStatus.New; break;
                case "active": parsed = ClashResultStatus.Active; break;
                case "reviewed": parsed = ClashResultStatus.Reviewed; break;
                case "approved": parsed = ClashResultStatus.Approved; break;
                case "resolved": parsed = ClashResultStatus.Resolved; break;
                default:
                    throw new ArgumentException(
                        $"unknown status '{statusStr}'. " +
                        "Valid: New, Active, Reviewed, Approved, Resolved.");
            }

            DocumentClash dc = doc.GetClash();

            // ── Locate parent test ─────────────────────────────────────────
            SavedItemReference testRef = null;
            foreach (SavedItem si in dc.TestsData.Value.TestsRoot.Children)
            {
                if (si is ClashTest t &&
                    string.Equals(t.DisplayName, testName, StringComparison.OrdinalIgnoreCase))
                {
                    testRef = dc.TestsData.CreateReference(t);
                    break;
                }
            }
            if (testRef == null)
            {
                throw new InvalidOperationException($"clash test '{testName}' not found");
            }

            // ── Locate target group ────────────────────────────────────────
            var freshTest = (ClashTest)dc.TestsData.ResolveReference(testRef);
            SavedItemReference groupRef = null;
            foreach (SavedItem child in freshTest.Children)
            {
                if (child is ClashResultGroup g &&
                    string.Equals(g.DisplayName, groupName, StringComparison.OrdinalIgnoreCase))
                {
                    groupRef = dc.TestsData.CreateReference(g);
                    break;
                }
            }
            if (groupRef == null)
            {
                throw new InvalidOperationException(
                    $"clash group '{groupName}' not found in test '{testName}'");
            }

            // ── Apply status ───────────────────────────────────────────────
            //  Assignee is required by the API for audit purposes. The default
            //  constructor yields an empty/anonymous assignee, which is fine
            //  for automation. If you want to tag changes with a specific name
            //  later, swap to `new Assignee("MCP Bridge")` or similar.
            var assignee = new Assignee();

            var freshGroup = (ClashResultGroup)dc.TestsData.ResolveReference(groupRef);

            string method;
            int childCount = 0;

            if (freshGroup is IClashResult groupAsResult)
            {
                // ClashResultGroup implements IClashResult: direct path.
                dc.TestsData.TestsEditResultStatus(groupAsResult, parsed, assignee);
                method = "group-direct";
            }
            else
            {
                // Fallback: iterate children and set each. Collect refs first
                // to avoid weakref issues during mutation.
                var childRefs = new List<SavedItemReference>();
                foreach (SavedItem gc in freshGroup.Children)
                {
                    if (gc is ClashResult cr)
                    {
                        childRefs.Add(dc.TestsData.CreateReference(cr));
                    }
                }

                foreach (var cRef in childRefs)
                {
                    var freshChild = dc.TestsData.ResolveReference(cRef);
                    if (freshChild is IClashResult childAsResult)
                    {
                        dc.TestsData.TestsEditResultStatus(childAsResult, parsed, assignee);
                        childCount++;
                    }
                }

                method = "children-iterated";
            }

            return new
            {
                test_name = testName,
                group_name = groupName,
                status = parsed.ToString(),
                method = method,
                children_updated = childCount
            };
        }



    }
}

/*
================================================================================
  EDITS TO MAKE IN YOUR EXISTING main .cs FILE
================================================================================

  EDIT 1 — make Handlers partial so this file can extend it.

      Find:
            internal static class Handlers
      Replace with:
            internal static partial class Handlers


  EDIT 2 — add three routes inside the Dispatch() switch in MCPBridgePlugin.

      Right after the line:
            case "/clash/delete":           return Handlers.ClashDelete(body);

      Insert:
            case "/clash/group/create":     return Handlers.ClashGroupCreate(body);
            case "/clash/group/list":       return Handlers.ClashGroupList(body);
            case "/clash/group/delete":     return Handlers.ClashGroupDelete(body);


  EDIT 3 (optional, cosmetic) — bump the route count from 35 to 38
      in the file-header comment and in the /health response object.

================================================================================
  PYTHON MCP SERVER TOOLS to add to server_http.py
================================================================================

    @mcp.tool()
    def create_clash_group(test_name: str, group_name: str) -> dict:
        """Create an (empty) clash group inside an existing clash test."""
        return _post("/clash/group/create", {
            "test_name":  test_name,
            "group_name": group_name,
        })

    @mcp.tool()
    def list_clash_groups(test_name: str) -> dict:
        """List clash groups inside a test, with their child clash names."""
        return _post("/clash/group/list", {"test_name": test_name})

    @mcp.tool()
    def delete_clash_group(test_name: str, group_name: str) -> dict:
        """Delete a clash group (and its child results) from a test."""
        return _post("/clash/group/delete", {
            "test_name":  test_name,
            "group_name": group_name,
        })
================================================================================
*/
