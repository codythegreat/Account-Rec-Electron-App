const inputExcelBtn = document.getElementById('input-excel-btn');
const copyTableBtn = document.getElementById('copy-table-btn');
const table = document.getElementById('table');
const {dialog} = require('electron').remote;
const Excel = require('exceljs');

var entries = [
]

inputExcelBtn.addEventListener("click", () => {

    dialog.showOpenDialog({
        properties: ['openFile'],
        filters: [
            { name: 'Spreadsheet', extensions: ['xlsx'] }
          ]
    }).then(result => {
            var workbook = new Excel.Workbook();
            workbook.xlsx.readFile(result.filePaths[0])
                .then(function() {
                    // get the default worksheet
                    var worksheet = workbook.getWorksheet('Sheet1');
                    var rowNumber = 1;

                    // append each line until no batch number is found
                    while (worksheet.getCell('R' + rowNumber).value !== null) {
                        entries.push({
                            desc: worksheet.getCell('G' + rowNumber).value,
                            amt: worksheet.getCell('H' + rowNumber).value,
                            date: worksheet.getCell('F' + rowNumber).value,
                            batch: worksheet.getCell('R' + rowNumber).value
                        })
                        rowNumber++;
                    }

                    // search for offsetting amounts, and zero them out
                    for (i=0;i<entries.length;i++) {
                        for (j=1;j<entries.length;j++) {
                            if (j==i) {
                                continue;
                            } else if (entries[i].amt+entries[j].amt<0.005&&entries[i].amt+entries[j].amt>-0.005) {
                                entries.splice(i, 1, {desc: null, amt: 0, date: null, batch: null});
                                entries.splice(j, 1, {desc: null, amt: 0, date: null, batch: null});
                            }
                        }
                    }

                    // for every item that has an amount, print it onto the table
                    for (i=0;i<entries.length;i++) {
                        if (entries[i].amt===0) {
                            continue;
                        } else {
                            var row = document.createElement("tr");

                            var desc = document.createElement("td");
                            desc.setAttribute('class', 'desc');
                            desc.innerHTML = entries[i].desc;
                            row.appendChild(desc);

                            var amt = document.createElement("td");
                            amt.setAttribute('class', 'amt');
                            amt.innerHTML = entries[i].amt;
                            row.appendChild(amt);

                            var date = document.createElement("td");
                            date.setAttribute('class', 'date');
                            date.innerHTML = entries[i].date.toLocaleDateString('en-US');
                            row.appendChild(date);

                            var batch = document.createElement("td");
                            batch.setAttribute('class', 'batch');
                            batch.innerHTML = entries[i].batch;
                            row.appendChild(batch);

                            table.getElementsByTagName('tbody')[0].appendChild(row);
                        }
                    }

                    document.getElementById('sum-text').innerHTML = "Sum: " + entries.reduce((a, b) => +a + +b.amt, 0);
                    
                });            
        }).catch(err => console.log(err));
});

copyTableBtn.addEventListener("click", () => {
    var body = document.body, range, sel;
    var tbody = document.getElementById("table").getElementsByTagName('tbody')[0];
    if (document.createRange && window.getSelection) {
        range = document.createRange();
        sel = window.getSelection();
        sel.removeAllRanges();
        range.selectNodeContents(tbody);
        sel.addRange(range);
    }
    document.execCommand("Copy");
});