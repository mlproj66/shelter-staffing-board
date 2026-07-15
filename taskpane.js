let assignments = [];

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

async function registerWorkbookEvents() {
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("tblAssignments");
        table.onChanged.add(async () => {
            await refreshBoard();
        });
        await context.sync();
    });
}

async function refreshBoard() {
    assignments = await loadAssignments();
    const shelters = groupAssignments(assignments);
    renderBoard(shelters);
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

function groupAssignments(rows) {
    const shelters = {};
    rows.forEach(row => {
        const shelterId = row.ShelterID;
        if (!shelters[shelterId]) {
            shelters[shelterId] = {
                shelterId: shelterId,
                shelterName: row.ShelterName,
                people: []
            };
        }
        shelters[shelterId].people.push(row);
    });
    return Object.values(shelters);
}

function renderBoard(shelters) {
    const board = document.getElementById("board");
    board.innerHTML = "";
    if (shelters.length === 0) {
        board.innerHTML = '<p style="color:#666">No assignments found. Check that tblAssignments has data.</p>';
        return;
    }
    shelters.forEach(shelter => {
        const shelterDiv = document.createElement("div");
        shelterDiv.className = "shelter";
        shelterDiv.dataset.shelterId = shelter.shelterId;
        shelterDiv.innerHTML = '<div class="shelterHeader">' + shelter.shelterName + '</div>';
        shelter.people.forEach(person => {
            const note = createPersonCard(person);
            shelterDiv.appendChild(note);
        });
        registerDropZone(shelterDiv);
        board.appendChild(shelterDiv);
    });
}

function createPersonCard(person) {
    const card = document.createElement("div");
    card.className = "person";
    card.draggable = true;
    card.dataset.assignmentId = person.AssignmentID;
    card.innerHTML =
        '<strong>' + (person.PersonName || person.PersonID) + '</strong><br>' +
        (person.Role || "") + '<br>' +
        (person.Shift || "");
    card.addEventListener("dragstart", dragStart);
    return card;
}

function dragStart(event) {
    event.dataTransfer.setData("assignmentId", event.target.dataset.assignmentId);
}

function registerDropZone(div) {
    div.addEventListener("dragover", e => { e.preventDefault(); });
    div.addEventListener("drop", async e => {
        e.preventDefault();
        const assignmentId = e.dataTransfer.getData("assignmentId");
        const shelterId = div.dataset.shelterId;
        await moveAssignment(assignmentId, shelterId);
        await refreshBoard();
    });
}

async function moveAssignment(assignmentId, shelterId) {
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("tblAssignments");
        const headerRange = table.getHeaderRowRange();
        const body = table.getDataBodyRange();
        headerRange.load("values");
        body.load("values");
        await context.sync();

        const headers = headerRange.values[0];
        const shelterColIdx = headers.indexOf("ShelterID");
        if (shelterColIdx === -1) return;

        const rowIndex = body.values.findIndex(r => String(r[0]) === String(assignmentId));
        if (rowIndex < 0) return;

        body.getCell(rowIndex, shelterColIdx).values = [[shelterId]];
        await context.sync();
    });
}

