const XLSX = require("xlsx");
const FS = require("fs");
const workbook = XLSX.readFile("Data.xlsx");
const sheet_name_list = workbook.SheetNames;
console.log(sheet_name_list); // getting as Sheet1

sheet_name_list.forEach(function (y) {
  let worksheet = workbook.Sheets[y];
  //getting the complete sheet
  console.log(worksheet);

  let headers = {};
  let data = [];
  for (z in worksheet) {
    if (z[0] === "!") continue;
    //parse out the column, row, and value
    let col = z.substring(0, 1);
    console.log(col);

    let row = parseInt(z.substring(1));
    console.log(row);

    let value = worksheet[z].v;
    console.log(value);

    //store header names
    if (row == 1) {
      headers[col] = value;
      // storing the header names
      continue;
    }

    if (!data[row]) data[row] = {};
    data[row][headers[col]] = value;
  }
  //drop those first two rows which are empty
  data.shift();
  data.shift();
  console.log(data);

  FS.writeFileSync("data/" + y + ".json", JSON.stringify(data, null, 4));
});
