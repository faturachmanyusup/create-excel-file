let xl = require('excel4node');
let path = require('path');

let defaultStyle = {
  font: {
    color: "#000000",
    size: 11,
  },
  numberFormat: "$#,##0.00; ($#,##0.00); -",
};

module.exports = {
  create: async (headers, values, options) => {
    try {
      const wb = new xl.Workbook();
      const ws = wb.addWorksheet(options.worksheet || "Data");
      const style = wb.createStyle(options.style || defaultStyle);
  
      headers.map(async (header, index) => {
        let row = 2;
        let column = index + 1;
  
        await ws.cell(1, column).string(header.label).style(style);
  
        for (let instance of values) {
          let key = header.key;
          let value = instance[key];
  
          if (typeof value === 'boolean') {
            await ws.cell(row, column).bool(value).style(style).style({font: {size: 14}});
          } else if (!isNaN(value)) {
            await ws.cell(row, column).number(value).style(style);  
          } else {
            await ws.cell(row, column).string(value).style(style);
          }
  
          row += 1;
        }
      });
  
      const filename = options.filename + ".xlsx";
      const fileDir = path.join(options.folder, filename);
  
      await wb.write(fileDir);
  
      console.log(`successfully create file ${filename}`);
    } catch (err) {
      console.log(err, "<<<< ERROR")
        return {
          message: `failed create file ${filename}`,
          error: err
        };
    }
  }
}
