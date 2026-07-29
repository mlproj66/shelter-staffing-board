let assignments = [];
let people = [];
let peopleById = new Map();
let shelterList = [];

function nextAssignmentId() {
    return "A-" + crypto.randomUUID().slice(0, 8);
}

async function pruneBlankAssignmentRows() {
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("tblAssignments");
        const headerRange = table.getHeaderRowRange();
        const body = table.getDataBodyRange();
        headerRange.load("values");
        body.load("values");
        await context.sync();

        const headers = headerRange.values[0];
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

Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Excel) {
        document.getElementById("board").innerHTML =
            '<p style="color:red">This add-in requires Excel.</p>';
        return;
    }
    try {
        await refreshBoard();
        await registerWorkbookEvents();
    } catch (e) {
        document.getElementById("board").innerHTML =
            '<p style="color:red">Error: ' + e.message + '</p>';
    }
});

let refreshTimer = null;
async function registerWorkbookEvents() {
    await Excel.run(async (context) => {
        ["tblAssignments", "tblShelters", "tblPeople"].forEach(name => {
            context.workbook.tables.getItem(name).onChanged.add(async () => {
                clearTimeout(refreshTimer);
                refreshTimer = setTimeout(() => refreshBoard().catch(console.error), 800);
            });
        });
        await context.sync();
    });
}

async function refreshBoard() {
    assignments = await loadAssignments();
    people = await loadPeople();
    shelterList = await loadShelters();
    pruneBlankAssignmentRows().catch(console.error);
    peopleById = new Map(people.map(p => [String(p.PersonID), p]));
    renderFromState();
}

function renderFromState() {
    const assignedIds = new Set(assignments.map(a => String(a.PersonID)));
    const unassigned = people.filter(p => !assignedIds.has(String(p.PersonID)));
    const shelters = groupAssignments(assignments);
    renderBoard(shelters, unassigned);
}

async function loadShelters() {
    return Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("tblShelters");
        const headerRange = table.getHeaderRowRange();
        const bodyRange = table.getDataBodyRange();
        headerRange.load("values");
        bodyRange.load("values");
        await context.sync();

        const headers = headerRange.values[0];
        return bodyRange.values
            .filter(r => r[0] !== "" && r[0] !== null)
            .map(row => {
                const obj = {};
                headers.forEach((h, i) => { obj[h] = row[i]; });
                return obj;
            });
    });
}

async function loadAssignments() {
    return Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("tblAssignments");
        const headerRange = table.getHeaderRowRange();
        const bodyRange = table.getDataBodyRange();
        headerRange.load("values");
        bodyRange.load("values");
        await context.sync();

        const headers = headerRange.values[0];
        const rows = bodyRange.values;
        return rows
            .filter(r => r[0] !== "" && r[0] !== null)
            .map(row => {
                const obj = {};
                headers.forEach((header, i) => { obj[header] = row[i]; });
                return obj;
            });
    });
}

async function loadPeople() {
    return Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("tblPeople");
        const headerRange = table.getHeaderRowRange();
        const bodyRange = table.getDataBodyRange();
        headerRange.load("values");
        bodyRange.load("values");
        await context.sync();

        const headers = headerRange.values[0];
        return bodyRange.values
            .filter(r => r[0] !== "" && r[0] !== null)
            .map(row => {
                const obj = {};
                headers.forEach((h, i) => { obj[h] = row[i]; });
                return obj;
            });
    });
}

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

function roleRank(role) {
    const r = String(role || "").trim().toUpperCase();
    if (r === "SV" || r.startsWith("SV ") || r.startsWith("SV-")) return 1;
    if (r === "SA" || r.startsWith("SA ") || r.startsWith("SA-")) return 2;
    return 3;
}

function renderBoard(shelters, unassigned) {
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
            col.dataset.shelterId = shelter.shelterId;
            col.dataset.shelterName = shelter.shelterName;
            col.dataset.shift = shift;
            col.innerHTML = '<div class="shiftHeader">' + shift + '</div>';

            const people = shelter.people
                .filter(p => p.Shift === shift)
                .sort((a, b) =>
                    roleRank(a.Role) - roleRank(b.Role) ||
                    String(a.PersonName).localeCompare(String(b.PersonName)));

            people.forEach(person => col.appendChild(createPersonCard(person)));
            registerDropZone(col);
            columns.appendChild(col);
        });

        shelterDiv.appendChild(columns);
        board.appendChild(shelterDiv);
    });
}

function createPersonCard(person) {
    const card = document.createElement("div");
    card.className = "person";
    card.draggable = true;
    card.dataset.assignmentId = person.AssignmentID;
    const p = peopleById.get(String(person.PersonID));
    const name = (p && p.Name) || person.PersonName || person.PersonID;
    card.innerHTML =
        '<strong>' + name + '</strong><br>' +
        (person.Role || "");
    card.addEventListener("dragstart", dragStart);
    return card;
}

function dragStart(event) {
    event.dataTransfer.setData("assignmentId", event.currentTarget.dataset.assignmentId);
}

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
            showRolePicker(e.pageX, e.pageY, (role) => {
                const newId = nextAssignmentId();
                assignments.push({
                    AssignmentID: newId, ShelterID: shelterId, PersonID: personId,
                    Shift: shift, Role: role, ShelterName: shelterName,
                    PersonName: (peopleById.get(String(personId)) || {}).Name || personId
                });
                renderFromState();
                createAssignment(newId, personId, shelterId, shift, role)
                    .catch(err => { console.error(err); refreshBoard(); });
            });
        } else if (assignmentId) {
            const a = assignments.find(x => String(x.AssignmentID) === String(assignmentId));
            if (a) {
                a.ShelterID = shelterId; a.ShelterName = shelterName; a.Shift = shift;
                renderFromState();
            }
            moveAssignment(assignmentId, shelterId, shift)
                .catch(err => { console.error(err); refreshBoard(); });
        }
    });
}

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
            deleteAssignment(assignmentId)
                .catch(err => { console.error(err); refreshBoard(); });
        }
    });
}

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

    setTimeout(() => {
        document.addEventListener("click", function dismiss(ev) {
            if (!picker.contains(ev.target)) {
                picker.hidden = true;
                document.removeEventListener("click", dismiss);
            }
        });
    }, 0);
}

async function createAssignment(newId, personId, shelterId, shift, role) {
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("tblAssignments");
        const headerRange = table.getHeaderRowRange();
        const body = table.getDataBodyRange();
        headerRange.load("values");
        body.load("values");
        await context.sync();

        const headers = headerRange.values[0];

        const newRow = headers.map(h => {
            if (h === "AssignmentID") return newId;
            if (h === "ShelterID") return shelterId;
            if (h === "PersonID") return personId;
            if (h === "Shift") return shift;
            if (h === "Role") return role;
            return "";
        });
        table.rows.add(null, [newRow]);
        await context.sync();

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

async function moveAssignment(assignmentId, shelterId, shift) {
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("tblAssignments");
        const headerRange = table.getHeaderRowRange();
        const body = table.getDataBodyRange();
        headerRange.load("values");
        body.load("values");
        await context.sync();

        const headers = headerRange.values[0];
        const shelterColIdx = headers.indexOf("ShelterID");
        const shiftColIdx = headers.indexOf("Shift");
        if (shelterColIdx === -1 || shiftColIdx === -1) return;

        const rowIndex = body.values.findIndex(r => String(r[0]) === String(assignmentId));
        if (rowIndex < 0) return;

        body.getCell(rowIndex, shelterColIdx).values = [[shelterId]];
        body.getCell(rowIndex, shiftColIdx).values = [[shift]];
        await context.sync();
    });
}

async function deleteAssignment(assignmentId) {
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("tblAssignments");
        const body = table.getDataBodyRange();
        body.load("values");
        await context.sync();
        const rowIndex = body.values.findIndex(r => String(r[0]) === String(assignmentId));
        if (rowIndex < 0) return;
        table.rows.getItemAt(rowIndex).delete();
        await context.sync();
    });
}
