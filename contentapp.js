/**
 * contentapp.js
 *
 * Excel add-in contentapp for a shelter staffing board.
 *
 * Reads three Excel tables (tblAssignments, tblPeople, tblShelters),
 * renders a drag-and-drop board of shelters/shifts, and writes changes
 * (create/move/delete assignment) back to the tblAssignments table.
 * Also listens for external changes to those tables and re-renders.
 */

// ---------------------------------------------------------------------------
// In-memory state, kept in sync with the Excel tables via loadAll()/refreshBoard()
// ---------------------------------------------------------------------------

let assignments = [];      // Rows from tblAssignments, as plain objects keyed by header name
let people = [];           // Rows from tblPeople, as plain objects keyed by header name
let peopleById = new Map(); // PersonID (string) -> person object, for quick lookups
let shelterList = [];      // Rows from tblShelters, as plain objects keyed by header name
let assignHeaders = null;  // Cached header row of tblAssignments (column order/names)

/**
 * Generates a short, unique-ish assignment ID (e.g. "A-1a2b3c4d").
 */
function nextAssignmentId() {
    return "A-" + crypto.randomUUID().slice(0, 8);
}

// ---------------------------------------------------------------------------
// Add-in bootstrap
// ---------------------------------------------------------------------------

Office.onReady(async (info) => {
    // This add-in only makes sense inside Excel; bail out with a message otherwise.
    if (info.host !== Office.HostType.Excel) {
        document.getElementById("board").innerHTML =
            '<p style="color:red">This add-in requires Excel.</p>';
        return;
    }
    try {
        await refreshBoard();              // Initial load + render
        await registerWorkbookEvents();    // Watch for external table edits
        // Clean up any fully-blank rows left behind by Excel table row deletion,
        // without blocking startup on it.
        pruneBlankAssignmentRows().catch(console.error);
    } catch (e) {
        document.getElementById("board").innerHTML =
            '<p style="color:red">Error: ' + e.message + '</p>';
    }
});

// ---------------------------------------------------------------------------
// Workbook change detection
// ---------------------------------------------------------------------------

let refreshTimer = null;  // Debounce timer for onChanged-triggered refreshes
let suppressUntil = 0;    // Timestamp until which refreshes are treated as "echoes" of our own writes

/**
 * Marks that we just wrote to the workbook ourselves. Used to lengthen the
 * debounce delay briefly afterward, so our own edits don't cause a jarring
 * immediate re-render on top of the optimistic UI update.
 */
function noteLocalWrite() { suppressUntil = Date.now() + 2000; }

/**
 * Subscribes to onChanged events for all three tables. Any change
 * (from this add-in or elsewhere, e.g. a user editing cells directly)
 * triggers a debounced full refresh of the board.
 */
async function registerWorkbookEvents() {
    await Excel.run(async (context) => {
        ["tblAssignments", "tblShelters", "tblPeople"].forEach(name => {
            context.workbook.tables.getItem(name).onChanged.add(async () => {
                // Use a longer delay right after our own writes to avoid
                // refreshing mid-write; shorter delay otherwise for responsiveness.
                const delay = Date.now() < suppressUntil ? 2500 : 800;
                clearTimeout(refreshTimer);
                refreshTimer = setTimeout(() => refreshBoard().catch(console.error), delay);
            });
        });
        await context.sync();
    });
}

// ---------------------------------------------------------------------------
// Data loading helpers
// ---------------------------------------------------------------------------

/**
 * Converts a 2D array of table rows into an array of objects keyed by header name.
 * Skips rows whose first cell is empty/null (treated as blank/placeholder rows).
 */
function rowsToObjects(headers, rows) {
    return rows
        .filter(r => r[0] !== "" && r[0] !== null)
        .map(row => {
            const obj = {};
            headers.forEach((h, i) => { obj[h] = row[i]; });
            return obj;
        });
}

/**
 * Loads header + data rows for all three tables in a single Excel.run batch,
 * converting each to an array of plain objects. Also caches the assignments
 * table header row (assignHeaders) for later column-index lookups.
 */
async function loadAll() {
    return Excel.run(async (context) => {
        const t = context.workbook.tables;
        const aH = t.getItem("tblAssignments").getHeaderRowRange().load("values");
        const aB = t.getItem("tblAssignments").getDataBodyRange().load("values");
        const pH = t.getItem("tblPeople").getHeaderRowRange().load("values");
        const pB = t.getItem("tblPeople").getDataBodyRange().load("values");
        const sH = t.getItem("tblShelters").getHeaderRowRange().load("values");
        const sB = t.getItem("tblShelters").getDataBodyRange().load("values");
        await context.sync();

        assignHeaders = aH.values[0];
        return {
            assignments: rowsToObjects(aH.values[0], aB.values),
            people: rowsToObjects(pH.values[0], pB.values),
            shelters: rowsToObjects(sH.values[0], sB.values)
        };
    });
}

/**
 * Reloads all table data from Excel into in-memory state, rebuilds the
 * PersonID -> person lookup map, and re-renders the board.
 */
async function refreshBoard() {
    const data = await loadAll();
    assignments = data.assignments;
    people = data.people;
    shelterList = data.shelters;
    peopleById = new Map(people.map(p => [String(p.PersonID), p]));
    renderFromState();
}

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------

/**
 * Derives the "unassigned" pool (people with no assignment row) and the
 * shelter groupings from current in-memory state, then renders the board.
 */
function renderFromState() {
    const assignedIds = new Set(assignments.map(a => String(a.PersonID)));
    const unassigned = people.filter(p => !assignedIds.has(String(p.PersonID)));
    const shelters = groupAssignments(assignments);
    renderBoard(shelters, unassigned);
}

/**
 * Groups assignment rows by ShelterID. Seeds the result with every known
 * shelter (even if empty) so empty shelters still show up on the board,
 * and falls back to creating an ad-hoc group for any ShelterID present in
 * the assignments but missing from tblShelters.
 */
function groupAssignments(rows) {
    const shelters = {};
    shelterList.forEach(s => {
        shelters[s.ShelterID] = {
            shelterId: s.ShelterID,
            shelterName: s.Name,
            people: []
        };
    });
    rows.forEach(row => {
        const shelterId = row.ShelterID;
        if (!shelters[shelterId]) {
            shelters[shelterId] = {
                shelterId: shelterId,
                shelterName: row.ShelterName || shelterId,
                people: []
            };
        }
        shelters[shelterId].people.push(row);
    });
    return Object.values(shelters);
}

/**
 * Sort priority for roles within a shift column: SV (supervisor) first,
 * then SA, then everything else (e.g. MN).
 */
function roleRank(role) {
    const r = String(role || "").trim().toUpperCase();
    if (r === "SV" || r.startsWith("SV ") || r.startsWith("SV-")) return 1;
    if (r === "SA" || r.startsWith("SA ") || r.startsWith("SA-")) return 2;
    return 3;
}

/**
 * Renders the full board: the "Unassigned" pool of draggable people chips,
 * plus one column-pair (Day/Night) per shelter, each populated with
 * person cards sorted by role then name.
 */
function renderBoard(shelters, unassigned) {
    // --- Unassigned pool ---
    const pool = document.getElementById("pool");
    pool.innerHTML = '<div class="poolHeader">Unassigned (' + unassigned.length + ')</div>';
    const chips = document.createElement("div");
    chips.className = "poolChips";
    unassigned.forEach(p => {
        const chip = document.createElement("div");
        chip.className = "chip";
        chip.draggable = true;
        chip.textContent = p.Name;
        chip.addEventListener("dragstart", e => {
            e.dataTransfer.setData("personId", String(p.PersonID));
        });
        chips.appendChild(chip);
    });
    pool.appendChild(chips);
    registerPoolDropZone(pool);

    // --- Shelter / shift board ---
    const board = document.getElementById("board");
    board.innerHTML = "";
    const SHIFTS = ["Day", "Night"];
    shelters.forEach(shelter => {
        const shelterDiv = document.createElement("div");
        shelterDiv.className = "shelter";
        shelterDiv.innerHTML = '<div class="shelterHeader">' + shelter.shelterName + '</div>';

        const columns = document.createElement("div");
        columns.className = "shiftColumns";

        SHIFTS.forEach(shift => {
            const col = document.createElement("div");
            col.className = "shiftCol";
            // Stash shelter/shift context on the column element so the drop
            // handler knows where a card was dropped.
            col.dataset.shelterId = shelter.shelterId;
            col.dataset.shelterName = shelter.shelterName;
            col.dataset.shift = shift;
            col.innerHTML = '<div class="shiftHeader">' + shift + '</div>';

            const shiftPeople = shelter.people
                .filter(p => p.Shift === shift)
                .sort((a, b) =>
                    roleRank(a.Role) - roleRank(b.Role) ||
                    String(a.PersonName).localeCompare(String(b.PersonName)));

            shiftPeople.forEach(person => col.appendChild(createPersonCard(person)));
            registerDropZone(col);
            columns.appendChild(col);
        });

        shelterDiv.appendChild(columns);
        board.appendChild(shelterDiv);
    });
}

/**
 * Builds a draggable card element for a single assignment row.
 * Prefers the live name from peopleById (tblPeople) over the cached
 * PersonName on the assignment row, falling back to PersonID if neither exists.
 */
function createPersonCard(person) {
    const card = document.createElement("div");
    const role = String(person.Role || "").toUpperCase();

    card.className = "person";
    
    if (role === "SV") {
        card.classList.add("role-sv");
    } else if (role === "SA") {
        card.classList.add("role-sa");
    } else if (role === "MN") {
        card.classList.add("role-mn");
    }
    card.draggable = true;
    card.dataset.assignmentId = person.AssignmentID;
    const p = peopleById.get(String(person.PersonID));
    const name = (p && p.Name) || person.PersonName || person.PersonID;
    const dayOff = formatDayOff(person.DayOff);
    card.innerHTML =
        '<strong>' + name + '</strong><br>' +
        (person.Role || "") +
        '<div class="dayoff">Off: ' + (dayOff || "—") + '</div>';
    card.addEventListener("dragstart", dragStart);
    card.addEventListener("click", (e) => {
        if (e.target.closest(".dayoff") || e.detail === 2) {
            openDayOffPicker(card, person.AssignmentID);
        }
    });
    return card;
}

/**
 * Excel stores dates as serial numbers when read via the API, so convert
 */
function formatDayOff(v) {
    if (v === "" || v === null || v === undefined) return "";
    if (typeof v === "number") {
        const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
        return (d.getUTCMonth() + 1) + "/" + d.getUTCDate() + "/" + d.getUTCFullYear();
    }
    return String(v);
}

function openDayOffPicker(card, assignmentId) {
    const existing = document.getElementById("dayOffInput");
    if (existing) existing.remove();
    const input = document.createElement("input");
    input.type = "date";
    input.id = "dayOffInput";
    input.style.position = "absolute";
    const r = card.getBoundingClientRect();
    input.style.left = (window.scrollX + r.left) + "px";
    input.style.top = (window.scrollY + r.bottom) + "px";
    input.style.zIndex = 20;
    document.body.appendChild(input);
    input.focus();
    if (input.showPicker) { try { input.showPicker(); } catch (e) {} }

    input.addEventListener("change", async () => {
        const val = input.value;               // "YYYY-MM-DD" or ""
        input.remove();
        const a = assignments.find(x => String(x.AssignmentID) === String(assignmentId));
        if (a) { a.DayOff = val ? excelSerialFromISO(val) : ""; renderFromState(); }
        noteLocalWrite();
        setDayOff(assignmentId, val).catch(err => { console.error(err); refreshBoard(); });
    });
    input.addEventListener("blur", () => setTimeout(() => input.remove(), 150));
}

function excelSerialFromISO(iso) {
    const [y, m, d] = iso.split("-").map(Number);
    return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
}

async function setDayOff(assignmentId, isoDate) {
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("tblAssignments");
        const body = table.getDataBodyRange();
        body.load("values");
        await context.sync();

        const headers = assignHeaders;
        const idIdx = headers.indexOf("AssignmentID");
        const offIdx = headers.indexOf("DayOff");
        if (offIdx === -1) return;

        const rowIndex = body.values.findIndex(r => String(r[idIdx]) === String(assignmentId));
        if (rowIndex < 0) return;

        body.getCell(rowIndex, offIdx).values = [[isoDate || ""]];
        await context.sync();
    });
}


/**
 * Drag handler for existing assignment cards: stashes the assignment ID
 * so the drop target knows which assignment is being moved.
 */
function dragStart(event) {
    event.dataTransfer.setData("assignmentId", event.currentTarget.dataset.assignmentId);
}

// ---------------------------------------------------------------------------
// Drag-and-drop targets
// ---------------------------------------------------------------------------

/**
 * Wires up a shift column as a drop target. Handles two drag sources:
 *  - a person chip from the Unassigned pool (personId) -> creates a new assignment
 *  - an existing assignment card (assignmentId) -> moves it to this shelter/shift
 * Updates local state optimistically, re-renders immediately, then persists
 * the change to Excel; on failure, falls back to a full refresh from Excel.
 */
function registerDropZone(div) {
    div.addEventListener("dragover", e => {
        e.preventDefault();
        div.classList.add("dragover");
    });
    div.addEventListener("dragleave", () => div.classList.remove("dragover"));
    div.addEventListener("drop", async e => {
        e.preventDefault();
        div.classList.remove("dragover");
        const personId = e.dataTransfer.getData("personId");
        const assignmentId = e.dataTransfer.getData("assignmentId");
        const { shelterId, shelterName, shift } = div.dataset;

        if (personId) {
            // New assignment: ask for a role first, then create it.
            showRolePicker(e.pageX, e.pageY, (role) => {
                const newId = nextAssignmentId();
                assignments.push({
                    AssignmentID: newId, ShelterID: shelterId, PersonID: personId,
                    Shift: shift, Role: role, ShelterName: shelterName,
                    PersonName: (peopleById.get(String(personId)) || {}).Name || personId
                });
                renderFromState();
                noteLocalWrite();
                createAssignment(newId, personId, shelterId, shift, role)
                    .catch(err => { console.error(err); refreshBoard(); });
            });
        } else if (assignmentId) {
            // Existing assignment moved to a different shelter/shift.
            const a = assignments.find(x => String(x.AssignmentID) === String(assignmentId));
            if (a) {
                a.ShelterID = shelterId; a.ShelterName = shelterName; a.Shift = shift;
                renderFromState();
            }
            noteLocalWrite();
            moveAssignment(assignmentId, shelterId, shift)
                .catch(err => { console.error(err); refreshBoard(); });
        }
    });
}

/**
 * Wires up the "Unassigned" pool as a drop target: dropping an existing
 * assignment card here removes that assignment (returns the person to the pool).
 * Registration is guarded so this only runs once even though renderBoard()
 * rebuilds the pool's inner content on every render.
 */
function registerPoolDropZone(pool) {
    if (pool.dataset.dzRegistered) return;
    pool.dataset.dzRegistered = "1";
    pool.addEventListener("dragover", e => { e.preventDefault(); pool.classList.add("dragover"); });
    pool.addEventListener("dragleave", () => pool.classList.remove("dragover"));
    pool.addEventListener("drop", e => {
        e.preventDefault();
        pool.classList.remove("dragover");
        const assignmentId = e.dataTransfer.getData("assignmentId");
        if (assignmentId) {
            assignments = assignments.filter(a => String(a.AssignmentID) !== String(assignmentId));
            renderFromState();
            noteLocalWrite();
            deleteAssignment(assignmentId)
                .catch(err => { console.error(err); refreshBoard(); });
        }
    });
}

// ---------------------------------------------------------------------------
// Small popup UI for choosing a role when creating an assignment
// ---------------------------------------------------------------------------

/**
 * Shows a small floating menu of role buttons (SV/SA/MN) at the given
 * page coordinates. Calls onPick(role) when a role is chosen, and
 * dismisses itself on any click outside the picker.
 */
function showRolePicker(x, y, onPick) {
    const picker = document.getElementById("rolePicker");
    picker.innerHTML = '<div class="rpTitle">Role:</div>';
    ["SV", "SA", "MN"].forEach(role => {
        const btn = document.createElement("button");
        btn.textContent = role;
        btn.addEventListener("click", async () => {
            picker.hidden = true;
            await onPick(role);
        });
        picker.appendChild(btn);
    });
    picker.style.left = Math.min(x, window.innerWidth - 180) + "px";
    picker.style.top = y + "px";
    picker.hidden = false;

    // Defer attaching the outside-click listener so the click that opened
    // the picker doesn't immediately close it.
    setTimeout(() => {
        document.addEventListener("click", function dismiss(ev) {
            if (!picker.contains(ev.target)) {
                picker.hidden = true;
                document.removeEventListener("click", dismiss);
            }
        });
    }, 0);
}

// ---------------------------------------------------------------------------
// Excel write operations (persist local state changes back to tblAssignments)
// ---------------------------------------------------------------------------

/**
 * Appends a new row to tblAssignments for a freshly created assignment.
 * Builds the row in the table's actual column order, then fills in
 * XLOOKUP formulas for PersonName/ShelterName (if those columns exist)
 * so display names stay in sync with tblPeople/tblShelters.
 */
async function createAssignment(newId, personId, shelterId, shift, role) {
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("tblAssignments");
        const headers = assignHeaders;

        const newRow = headers.map(h => {
            if (h === "AssignmentID") return newId;
            if (h === "ShelterID") return shelterId;
            if (h === "PersonID") return personId;
            if (h === "Shift") return shift;
            if (h === "Role") return role;
            if (h === "DayOff") return "";
            return "";
        });
        table.rows.add(null, [newRow]);
        await context.sync();

        // Find the row we just added (it's the last row) and stamp in
        // lookup formulas for the derived name columns, if present.
        const body2 = table.getDataBodyRange();
        body2.load("rowCount");
        await context.sync();
        const last = body2.rowCount - 1;
        const pIdx = headers.indexOf("PersonName");
        const sIdx = headers.indexOf("ShelterName");
        if (pIdx > -1) body2.getCell(last, pIdx).formulas =
            [["=XLOOKUP([@PersonID], tblPeople[PersonID], tblPeople[Name], \"?\")"]];
        if (sIdx > -1) body2.getCell(last, sIdx).formulas =
            [["=XLOOKUP([@ShelterID], tblShelters[ShelterID], tblShelters[Name], \"?\")"]];
        await context.sync();
    });
}

/**
 * Updates an existing assignment row's ShelterID and Shift columns to
 * reflect a drag-and-drop move. Locates the row by matching AssignmentID.
 */
async function moveAssignment(assignmentId, shelterId, shift) {
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("tblAssignments");
        const body = table.getDataBodyRange();
        body.load("values");
        await context.sync();

        const headers = assignHeaders;
        const shelterColIdx = headers.indexOf("ShelterID");
        const shiftColIdx = headers.indexOf("Shift");
        if (shelterColIdx === -1 || shiftColIdx === -1) return;

        const idIdx = headers.indexOf("AssignmentID");
        const rowIndex = body.values.findIndex(r => String(r[idIdx]) === String(assignmentId));
        if (rowIndex < 0) return;

        body.getCell(rowIndex, shelterColIdx).values = [[shelterId]];
        body.getCell(rowIndex, shiftColIdx).values = [[shift]];
        await context.sync();
    });
}

/**
 * Deletes the tblAssignments row matching the given AssignmentID
 * (used when a card is dropped back into the Unassigned pool).
 */
async function deleteAssignment(assignmentId) {
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("tblAssignments");
        const body = table.getDataBodyRange();
        body.load("values");
        await context.sync();
        const assignmentColIdx = assignHeaders.indexOf("AssignmentID");
        const rowIndex = body.values.findIndex( r => String(r[assignmentColIdx]) === String(assignmentId));
        if (rowIndex < 0) return;
        table.rows.getItemAt(rowIndex).delete();
        await context.sync();
    });
}

/**
 * Housekeeping pass run at startup: removes any tblAssignments rows where
 * all of the key columns (AssignmentID, ShelterID, PersonID, Shift, Role)
 * are blank. These can be left behind by certain Excel table row-deletion
 * paths. Rows are deleted from bottom to top so earlier indices stay valid.
 */
async function pruneBlankAssignmentRows() {
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("tblAssignments");
        const body = table.getDataBodyRange();
        body.load("values");
        await context.sync();

        const headers = assignHeaders;
        const idx = ["AssignmentID", "ShelterID", "PersonID", "Shift", "Role"]
            .map(h => headers.indexOf(h))
            .filter(i => i > -1);

        const blank = [];
        body.values.forEach((r, i) => {
            if (idx.every(c => r[c] === "" || r[c] === null)) blank.push(i);
        });
        for (let i = blank.length - 1; i >= 0; i--) {
            table.rows.getItemAt(blank[i]).delete();
        }
        await context.sync();
    });
}
