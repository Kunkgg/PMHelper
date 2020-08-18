const xlsxFile = require("read-excel-file/node");
const path = require("path");

function chooseFile(name) {
  var chooser = document.querySelector(name);
  chooser.addEventListener(
    "change",
    function (evt) {
      console.log(this.value);
      drawChart(this.value);
    },
    false
  );

  chooser.click();
}
chooseFile("#fileDialog");

function drawChart(fpath) {
  // console.log(fpath);
  const title = path.basename(fpath, ".xlsx");

  xlsxFile(fpath, { getSheets: true }).then((sheets) => {
    xlsxFile(fpath, { sheet: sheets[0].name }).then((rows) => {
      var datas = formatData(rows);
      console.log(datas);

      Highcharts.ganttChart("container", {
        title: {
          text: title,
        },
        series: [
          {
            name: title,
            data: datas,
          },
        ],
      });
    });
  });
}

function formatData(rows) {
  var header = rows[0];
  console.log(header);
  var datas = [];

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
