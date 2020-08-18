const xlsxFile = require("read-excel-file/node");

const fpath = "../schedule.xlsx";

xlsxFile(fpath, { getSheets: true })
  .then((sheets) => {
    sheets.forEach((obj) => {
      console.log(obj.name);
    });

    return sheets[0].name;
  })
  .then((sheet) => {
    xlsxFile(fpath, { sheet: sheet }).then((rows) => {
      console.log(rows);
    });
  });

// xlsxFile("./Data.xlsx", { sheet: "Exec" }).then((rows) => {
//   for (i in rows) {
//     for (j in rows[i]) {
//       console.log(rows[i][j]);
//     }
//   }
// });
