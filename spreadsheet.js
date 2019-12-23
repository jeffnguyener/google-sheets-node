const GoogleSpreadsheet = require("google-spreadsheet");
const { promisify } = require("util");

const creds = require("./client_secret.json");

function printStudent(student) {
  console.log(`Name: ${student.studentname}`);
  console.log(`Major: ${student.major}`);
  console.log(`Home State: ${student.homestate}`);
  console.log("---------------");
}

async function accessSpreadsheet() {
  const doc = new GoogleSpreadsheet(
    "1A2azG69p0qvZx-9XHnX40GGLT9m_gCnYdqbrCdAkbDM"
  );
  await promisify(doc.useServiceAccountAuth)(creds);
  const info = await promisify(doc.getInfo)();
  const sheet = info.worksheets[0];
  //   console.log(`Title: ${sheet.title}, Rows: ${sheet.rowCount}`);

  //PULLING INFO FROM CERTAIN ROWS & COLUMNS
  const cells = await promisify(sheet.getCells)({
    "min-row": 1,
    "max-row": 3,
    "min-col": 1,
    "max-col": 2
  });

  for (const cell of cells) {
    console.log(`${cell.row}, ${cell.col}: ${cell.value}`);

    cells[2].value = "Alexis";
    cells[2].save();
  }

  //DELETES ROW WITH QUERY OF STUDENT NAME
  //   const rows = await promisify(sheet.getRows)({
  //     query: "studentname = Brent"
  //   });
  //   rows[0].del();

  //ADDING STUDENT ONTO NEW ROW
  // const row = {
  //     studentname: 'Brent',
  //     major: 'Computer Engineering',
  //     homestate: 'PA',
  //     classlevel: '5. Graduated',
  //     extracurricularactivity: 'Music'
  // }
  // await promisify(sheet.addRow)(row);

  //QUERYING AND MAKING CHANGES TO ROW
  //   const rows = await promisify(sheet.getRows)({
  //     query: 'homestate = NY'
  //     // offset: 5,
  //     // limit: 10,
  //     // orderby: 'homestate'
  //   });

  // //   console.log(rows)
  //   rows.forEach( row => {
  //     //   printStudent(row)
  //     row.homestate = 'PA'
  //     row.save()
  //     //^^ DOING SO COMMITS THE STUDENTS FROM NY TO PA TO THE SPREADSHEET
  //   })
}

accessSpreadsheet();
