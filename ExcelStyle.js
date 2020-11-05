"use strict";

import { saveAs } from "file-saver";
import XLSX from "./xlsx";
import * as Util from "./util";

function Workbook() {
  if (!(this instanceof Workbook)) return new Workbook();
  this.SheetNames = [];
  this.Sheets = {};
}

class ExcelStyle {
  constructor() {
    if (!(this instanceof ExcelStyle)) return new ExcelStyle();

    this.XLSX = XLSX;
    this.sheets = [];
  }

  getCharCol(R, C) {
    return XLSX.utils.encode_cell({
      c: C,
      r: R,
    });
  }

  sheetFromArrayOfArrays(data, styles) {
    let ws = {};
    let range = {
      s: {
        c: 10000000,
        r: 10000000,
      },
      e: {
        c: 0,
        r: 0,
      },
    };
    let hasStyleRows = (styles.rows && styles.rows.length > 0);
    for (var R = 0; R != data.length; ++R) {
      for (var C = 0; C != data[R].length; ++C) {
        if (range.s.r > R) range.s.r = R;
        if (range.s.c > C) range.s.c = C;
        if (range.e.r < R) range.e.r = R;
        if (range.e.c < C) range.e.c = C;
        var cell = {
          v: data[R][C],
        };
        if (cell.v == null) {
          cell.v = "";
        }
        var cell_ref = XLSX.utils.encode_cell({
          c: C,
          r: R,
        });

        if (typeof cell.v === "number") cell.t = "n";
        else if (typeof cell.v === "boolean") cell.t = "b";
        else if (cell.v instanceof Date) {
          cell.t = "n";
          cell.z = XLSX.SSF._table[14];
          cell.v = Util.datenum(cell.v);
        } else cell.t = "s";

        // add style
        if (styles.all) {
          cell.s = styles.all;
        }

        if (styles.header && R === 0) {
          cell.s = styles.header;
        }

        if (hasStyleRows) {
          let record = styles.rows.find(key => {
            return (key[R]);
          });
          if (record) {
            cell.s = record[R];
          }
        }

        ws[cell_ref] = cell;
      }
    }
    if (range.s.c < 10000000) ws["!ref"] = XLSX.utils.encode_range(range);
    return ws;
  }

  addSheet(...args) {
    const arg = args[0];
    if (!arg.data || !Array.isArray(arg.data))
      return new Error("sheet data isn't exist");

    const name = arg.name || `sheet${this.sheets.length + 1}`;

    let header = {};
    if (arg.header) {
      if (typeof arg.header !== "object")
        return new Error("header isn't object");
      header = arg.header;
    } else {
      Object.keys(arg.data[0]).forEach((val) => {
        header[val] = val;
      });
    }

    this.sheets.push({
      name: name,
      data: arg.data,
      header: header,
      merges: arg.merges || [],
      autoWidth: (arg.autoWidth) ? true : false,
      styles: arg.styles || {}
    });

    return this;
  }

  download(wb, filename = "output_excel", bookType = "xlsx") {
    var wbout = XLSX.write(wb, {
      bookType: bookType,
      bookSST: false,
      type: "binary",
    });
    saveAs(
      new Blob([Util.s2ab(wbout)], {
        type: "application/octet-stream",
      }),
      `${filename}.${bookType}`
    );
  }

  output(filename, bookType) {
    let wb = new Workbook();

    this.sheets.forEach((sheet) => {
      let data = Util.toJsonFormat(Object.keys(sheet.header), sheet.data);
      data = [...data];
      data.unshift(Object.values(sheet.header));
      let ws = this.sheetFromArrayOfArrays(data, sheet.styles);

      if (sheet.merges.length > 0) {
        if (!ws["!merges"]) ws["!merges"] = [];
        sheet.merges.forEach((item) => {
          ws["!merges"].push(XLSX.utils.decode_range(item));
        });
      }

      // 設定表格寬度
      if (sheet.autoWidth) {
        const colWidth = data.map((row) =>
          row.map((val) => {
            if (val == null || val == undefined) {
              return {
                wch: 10,
              };
            } else if (val.toString().charCodeAt(0) > 255) {
              return {
                wch: val.toString().length * 2,
              };
            } else {
              return {
                wch: val.toString().length * 1.5,
              };
            }
          })
        );

        let result = colWidth[1];
        for (let i = 1; i < colWidth.length; i++) {
          for (let j = 0; j < colWidth[i].length; j++) {
            if (result[j]["wch"] < colWidth[i][j]["wch"]) {
              result[j]["wch"] = colWidth[i][j]["wch"];
            }
          }
        }
        ws["!cols"] = result;
      }

      /* add worksheet to workbook */
      wb.SheetNames.push(sheet.name);
      wb.Sheets[sheet.name] = ws;
    });

    this.download(wb, filename, bookType);
  }
}

export default ExcelStyle;
