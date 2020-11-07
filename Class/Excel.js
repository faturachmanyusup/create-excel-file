let xl = require('excel4node');
let path = require('path');

let defaultStyle = {
  font: {
    color: "#000000",
    size: 11,
  },
  numberFormat: "#,##0.00; (#,##0.00); -",
};

class Excel {
  constructor(data, options, headers) {
    this.headers = headers || this.getHeaders(data);
    this.style = options.style || defaultStyle;
    this.filename = options.filename + ".xlsx";
    this.worksheet = options.worksheet || "Data";
    this.values = data;
    this.folder = options.folder;
  }

  getHeaders(data) {
    return Object.keys(data[0]);
  }

  async create() {
    try {
      const wb = new xl.Workbook();
      const ws = wb.addWorksheet(this.worksheet);
      const style = wb.createStyle(this.style);
  
      this.headers.map(async (header, index) => {
        let row = 2;
        let column = index + 1;
        let columnHeader = header.label || header;
  
        await ws.cell(1, column).string(columnHeader).style(style);
  
        for (let instance of this.values) {
          let key = header.key || header;
          let value = String(instance[key]);

          await ws.cell(row, column).string(value).style(style);
  
          row += 1;
        }
      });
  
      const fileDir = path.join(this.folder, this.filename);
  
      await wb.write(fileDir);
  
      console.log(`successfully create file ${this.filename}`);
    } catch (err) {
      console.log(err, "<<<< ERROR")
      return {
        message: `failed create file ${this.filename}`,
        error: err
      };
    }
  } 
}

module.exports = Excel;