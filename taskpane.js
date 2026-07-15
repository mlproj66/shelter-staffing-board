let assignments = [];

Office.onReady(async (info) => {

    if (info.host !== Office.HostType.Excel) {
        return;
    }

    await refreshBoard();

    await registerWorkbookEvents();

});

async function registerWorkbookEvents() {

    await Excel.run(async (context) => {

        context.workbook.onChanged.add(async () => {

            await refreshBoard();

        });

        await context.sync();

    });

}

async function refreshBoard() {

    assignments = await loadAssignments();

    const shelters =
        groupAssignments(assignments);

    renderBoard(shelters);

}

async function loadAssignments() {

    return Excel.run(async (context) => {

        const table =
            context.workbook.tables.getItem(
                "tblAssignments"
            );

        const range =
            table.getRange();

        range.load("values");

        await context.sync();

        const values = range.values;

        const headers = values[0];

        const rows = values.slice(1);

        const results = [];

        rows.forEach(row => {

            const obj = {};

            headers.forEach((header, index) => {

                obj[header] = row[index];

            });

            results.push(obj);

        });

        return results;

    });

}

function groupAssignments(rows) {

    const shelters = {};

    rows.forEach(row => {

        const shelterId =
            row.ShelterID;

        if (!shelters[shelterId]) {

            shelters[shelterId] = {

                shelterId:
                    shelterId,

                shelterName:
                    row.ShelterName,

                people: []

            };

        }

        shelters[shelterId]
            .people
            .push(row);

    });

    return Object.values(shelters);

}

function renderBoard(shelters) {

    const board =
        document.getElementById("board");

    board.innerHTML = "";

    shelters.forEach(shelter => {

        const shelterDiv =
            document.createElement("div");

        shelterDiv.className =
            "shelter";

        shelterDiv.dataset.shelterId =
            shelter.shelterId;

        shelterDiv.innerHTML =
            `
            <div class="shelterHeader">
                ${shelter.shelterName}
            </div>
            `;

        shelter.people.forEach(person => {

            const note =
                createPersonCard(person);

            shelterDiv.appendChild(note);

        });

        registerDropZone(shelterDiv);

        board.appendChild(shelterDiv);

    });

}

function createPersonCard(person) {

    const card =
        document.createElement("div");

    card.className =
        "person";

    card.draggable = true;

    card.dataset.assignmentId =
        person.AssignmentID;

    const daysOff =
        person.DaysOff || "";

    card.innerHTML =
        `
        <strong>${person.PersonName}</strong><br>
        ${person.Role}<br>
        ${person.Start} - ${person.End}<br>
        Off: ${daysOff}
        `;

    card.addEventListener(
        "dragstart",
        dragStart
    );

    return card;

}

function dragStart(event) {

    event.dataTransfer.setData(
        "assignmentId",
        event.target.dataset.assignmentId
    );

}

function registerDropZone(div) {

    div.addEventListener(
        "dragover",
        e => {

            e.preventDefault();

        }
    );

    div.addEventListener(
        "drop",
        async e => {

            e.preventDefault();

            const assignmentId =
                e.dataTransfer.getData(
                    "assignmentId"
                );

            const shelterId =
                div.dataset.shelterId;

            await moveAssignment(
                assignmentId,
                shelterId
            );

            await refreshBoard();

        }
    );

}

async function moveAssignment(
    assignmentId,
    shelterId
) {

    await Excel.run(async (context) => {

        const table =
            context.workbook.tables.getItem(
                "tblAssignments"
            );

        const body =
            table.getDataBodyRange();

        body.load("values");

        await context.sync();

        const values =
            body.values;

        const rowIndex =
            values.findIndex(
                r => r[0] === assignmentId
            );

        if (rowIndex < 0)
            return;

        const shelterColumnIndex = 1;

        body.getCell(
            rowIndex,
            shelterColumnIndex
        ).values = [[shelterId]];

        await context.sync();

    });

}

