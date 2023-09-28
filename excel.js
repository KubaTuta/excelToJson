let selectedFile;
console.log(window.XLSX);
document.getElementById("input").addEventListener("change", (event) => {
  selectedFile = event.target.files[0];
});

let data = [
  {
    name: "jayanth",
    data: "scd",
    abc: "sdef",
  },
];

document.getElementById("button").addEventListener("click", () => {
  XLSX.utils.json_to_sheet(data, "out.xlsx");
  if (selectedFile) {
    let fileReader = new FileReader();
    fileReader.readAsBinaryString(selectedFile);
    fileReader.onload = (event) => {
      let data = event.target.result;
      let workbook = XLSX.read(data, { type: "binary" });
      console.log(workbook);
      let resultArray = [];

      workbook.SheetNames.forEach((sheet) => {
        let worksheet = workbook.Sheets[sheet];

        let range = XLSX.utils.decode_range(worksheet["!ref"]);
        let rows = range.e.r;

        for (let i = 1; i <= rows; i++) {
          let cellB = worksheet[XLSX.utils.encode_cell({ r: i, c: 0 })];
          let cellA = worksheet[XLSX.utils.encode_cell({ r: i, c: 1 })];
          if (cellA && cellB) {
            let key = cellA.v;
            let value = cellB.v;
            let obj = {};
            obj[key] = value;
            resultArray.push(obj);
          }
        }
      });

      console.log(resultArray);
      document.getElementById("jsondata").innerHTML = JSON.stringify(
        resultArray,
        undefined,
        4
      );
    };
  }
});
