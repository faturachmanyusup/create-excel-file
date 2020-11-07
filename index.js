const Excel = require('./Class/Excel');
const excel = require('./excel');

const headers = [
  {
    label: "ID",
    key: "id",
  },
  {
    label: "Name",
    key: "name",
  },
  {
    label: "Age",
    key: "age",
  },
  {
    label: "Favorite Color",
    key: "favColor",
  }
]

const json = [
  {
    id: 1,
    name: "John",
    age: 20,
    address: "Jakarta",
    favColor: "Red",
    active: true
  },
  {
    id: 2,
    name: "Doe",
    age: 30,
    address: "Bali",
    favColor: "Blue",
    active: true
  },
  {
    id: 3,
    name: "Jonathan",
    age: 19,
    address: "Silicon Valley",
    favColor: "White",
    active: false
  }
]

const options = {
  style: {
    font: {
      color: "#000000",
      size: 12,
    },
    alignment: {
      horizontal: "center"
    }
  },
  worksheet: "my_worksheet",
  filename: "my_file",
  folder: 'xlsx'
}

// create file using class with custom headers
const FileWithCustomHeaders = new Excel(json, {
  ...options,
  filename: "class_with_custom_headers"
}, headers);
FileWithCustomHeaders.create();


// create file using class without custom headers
const FileWithoutCustomHeaders = new Excel(json, {
  ...options,
  filename: "class_without_custom_headers"
});
FileWithoutCustomHeaders.create();


// create file using function must include headers
excel.create(headers, json, options);