const xlsxFile = require("read-excel-file/node");
var Highcharts = require("highcharts");
// Load module after Highcharts is loaded
require("highcharts/modules/gantt")(Highcharts);

const fpath = "../schedule.xlsx";

xlsxFile(fpath, { getSheets: true })
  .then((sheets) => {
    // sheets.forEach((obj) => {
    //   console.log(obj.name);
    // });

    xlsxFile(fpath, { sheet: sheets[0].name }).then((rows) => {
      // console.log(rows);
      // console.log(rows[0]);

      return formatData(rows);
      // console.log(datas);
    });
  })
  .then((datas) => {
    Highcharts.ganttChart("container", {
      title: {
        text: "综合布线施工进度计划",
      },

      series: [
        {
          name: "三星工厂综合布线",
          data: datas,
        },
      ],
    });
  });

function formatData(rows) {
  let header = rows[0];
  var datas = [];

  rows.slice(1).forEach((row) => {
    var data = {};

    for (i = 0; i < header.length; i++) {
      // console.log(header[i]);
      data[header[i]] = row[i];
    }
    datas.push(data);
  });

  return datas;
}
