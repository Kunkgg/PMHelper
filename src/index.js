const xlsxFile = require("read-excel-file/node");
// var Highcharts = require("highcharts");
// Load module after Highcharts is loaded
// require("highcharts/modules/gantt")(Highcharts);

const fpath = "../schedule.xlsx";
var datas = [];

xlsxFile(fpath, { getSheets: true }).then((sheets) => {
  // sheets.forEach((obj) => {
  //   console.log(obj.name);
  // });

  xlsxFile(fpath, { sheet: sheets[0].name }).then((rows) => {
    console.log(rows);
    console.log(rows[0]);
    console.log(rows[1]);
    console.log(rows[2]);

    var datas = formatData(rows);
    console.log(datas);

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
    return datas;
  });
});

function formatData(rows) {
  var header = rows[0];
  console.log(header);
  // var datas = [];

  rows.slice(1).forEach((row) => {
    var data = {};

    for (i = 0; i < header.length; i++) {
      let value = row[i];
      if (value) {
        if (header[i] == "start" || header[i] == "end") {
          data[header[i]] = formatDate(value);
        } else if (header[i] == "dependency") {
          data[header[i]] = formatDep(value.toString());
        } else if (header[i] == "id" || header[i] == "parent") {
          if (typeof value != "string") {
            data[header[i]] = value.toString();
          }
        } else {
          data[header[i]] = value;
        }
      }
    }
    datas.push(data);
  });

  return datas;
}

function formatDate(date) {
  return Date.UTC(date.getFullYear(), date.getMonth() + 1, date.getDate());
}

function formatDep(dependency) {
  return dependency.replace(/，/gi, ",").split(",");
}
