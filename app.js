const express = require('express');
const app = express();
// const fs = require('fs')
const bodyParser = require('body-parser');
var XLSX = require('xlsx')
// request = require('request');
const HTTP = require('http');
// var FileSaver = require('file-saver');
var path = require('path');
var cors = require('cors');
// var mime = require('mime');
// const Blob = require('cross-blob');
// const { request } = require('http');
// request = require('request');

const FormulaParser = require('hot-formula-parser').Parser;
const parser = new FormulaParser();

const Excel = require('exceljs');
// const { style } = require('@angular/animations');
const workbook = new Excel.Workbook();
const calcworkbook = new Excel.Workbook();

//const jsonParser = bodyParser.json();

app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());
app.use(cors());
// xlsxFile('./data/item_master.xlsx').then((rows) => {
// });



app.post('/casingCalc', function (req, res) {
  let calcdata = [];
  console.log(req.body);
  let wb = XLSX.readFile('./data/calc/casing.xlsx');
  let ws = wb.Sheets['Sheet2'];
  // console.log(supplyLength);


  // console.log('e'+ eDh +'  '+ 'd' + eDw);

  // var wopts = { bookType:'xlsx', bookSST:false, type:'file' };
  // XLSX.writeFile(wb, './data/calc/casing.xlsx', wopts);

  // const casingCalcData = XLSX.utils.sheet_to_json(wb);
  // console.log(casingCalcData);

  function getCellResult(worksheet, cellLabel) {
    if (worksheet.getCell(cellLabel).formula) {
      return parser.parse(worksheet.getCell(cellLabel).formula).result;
    } else {
      return worksheet.getCell(cellCoord.label).value;
    }
  }




  workbook.xlsx.readFile('./data/calc/bom_temp.xlsx').then(() => {

    calcworkbook.xlsx.readFile('./data/calc/casing.xlsx').then(() => {
      var worksheet = calcworkbook.getWorksheet(1);
      console.log(req.body);
      worksheet.getCell('A2').value = req.body.unitForm.innerSheet.Description;
      worksheet.getCell('B2').value = req.body.unitForm.outerSheet.Description;

      var supplyLength = { sl: req.body.unitForm.supplyDimension };
      var exhaustLength = { el: req.body.unitForm.exhaustDimension };
      const supplyLengthSum = Object.values(supplyLength).reduce((a, v) => a += v.reduce((a, ob) => a += ob.length, 0), 0);
      const exhaustLengthSum = Object.values(exhaustLength).reduce((a, v) => a += v.reduce((a, ob) => a += ob.length, 0), 0);

      worksheet.getCell('C2').value = supplyLengthSum;
      worksheet.getCell('D2').value = req.body.unitForm.supplyDimension[0].height;
      worksheet.getCell('E2').value = req.body.unitForm.supplyDimension[0].width;

      if (exhaustLengthSum == '00') {
        worksheet.getCell('G2').value = 00;
      } else { worksheet.getCell('G2').value = exhaustLengthSum; }
      // console.log(exhaustLengthSum);
      let eDh; let eDw;
      if (req.body.unitForm.exhaustDimension[0].height == "") {
        eDh = 00;
      } else { eDh = req.body.unitForm.exhaustDimension[0].height }
      if (req.body.unitForm.exhaustDimension[0].width == "") {
        eDw = 00;
      } else { eDw = req.body.unitForm.exhaustDimension[0].width }

      worksheet.getCell('H2').value = eDh;
      worksheet.getCell('I2').value = eDw;
      worksheet.getCell('F2').value = req.body.unitForm.supplyDimension.length;
      worksheet.getCell('J2').value = req.body.unitForm.exhaustDimension.length;
      if (req.body.unitForm.panelThick == '') { worksheet.getCell('K2').value = 00 } else {
        worksheet.getCell('K2').value = req.body.unitForm.panelThick;
      }

      calcworkbook.xlsx.writeFile('./data/calc/casing.xlsx');

      parser.on('callCellValue', function (cellCoord, done) {
        if (worksheet.getCell(cellCoord.label).formula) {
          done(parser.parse(worksheet.getCell(cellCoord.label).formula).result);
        } else {
          done(worksheet.getCell(cellCoord.label).value);
        }
      });

      parser.on('callRangeValue', function (startCellCoord, endCellCoord, done) {
        var fragment = [];

        for (var row = startCellCoord.row.index; row <= endCellCoord.row.index; row++) {
          var colFragment = [];

          for (var col = startCellCoord.column.index; col <= endCellCoord.column.index; col++) {
            colFragment.push(worksheet.getRow(row + 1).getCell(col + 1).value);
          }

          fragment.push(colFragment);
        }

        if (fragment) {
          done(fragment);
        }
      });

      var area = getCellResult(worksheet, 'L2');
      var isw = getCellResult(worksheet, 'M2');
      var osw = getCellResult(worksheet, 'N2');
      var O = getCellResult(worksheet, 'O2');
      var P = getCellResult(worksheet, 'P2');
      var Q = getCellResult(worksheet, 'Q2');
      var R = getCellResult(worksheet, 'R2');
      calcdata = XLSX.utils.sheet_to_json(ws);
      calcdata[0].area = area;
      calcdata[0].inner_sheet_weight = isw;
      calcdata[0].outer_sheet_weight = osw;
      calcdata[0].corner_profile = O;
      calcdata[0].omega_profile = P;
      calcdata[0].polyol = Q;
      calcdata[0].isol = R;
      console.log(calcdata);
      // res.send(calcdata);
      console.log(calcdata);
      // console.log(workbook);

      var ahuCasing = req.body.unitForm;
      let ahuCasingData = [];
      var innerSkin = {
        part_code: req.body.unitForm.innerSheet.Code,
        description: 'Casing Inner Sheet', specification: req.body.unitForm.innerSheet.Name, type: '',
        qty: calcdata[0].inner_sheet_weight, uom: req.body.unitForm.innerSheet.Unit, totalQty: calcdata[0].inner_sheet_weight * req.body.ahuQty
      }
      var outerSkin = {
        part_code: req.body.unitForm.outerSheet.Code,
        description: 'Casing Outer Sheet', specification: req.body.unitForm.outerSheet.Name, type: '',
        qty: calcdata[0].outer_sheet_weight, uom: req.body.unitForm.outerSheet.Unit, totalQty: calcdata[0].outer_sheet_weight * req.body.ahuQty
      }
      ahuCasingData.push(innerSkin, outerSkin);
      console.log(ahuCasingData);

      var sheet1 = workbook.getWorksheet(1);
      sheet1.columns
      console.log(sheet1.getCell('A1').value);
      sheet1.getCell('A1').value = 'Bill Of Materials';
      sheet1.getCell(3, 3).value = req.body.project;
      sheet1.getCell(4, 3).value = req.body.ahuType;
      sheet1.getCell(5, 3).value = req.body.unitForm.airVolume;
      sheet1.getCell('C6').value = req.body.ahuQty;
      sheet1.getCell('C7').value = req.body.ahuModel;
      sheet1.getCell('D6').value = 'DATE                  :    ' + (new Date());
      sheet1.mergeCells('A9:H9');
      sheet1.getCell('A9').alignment = { horizontal: 'center', vertical: 'middle' };
      // sheet1.getCell('A9').style = {fill: {bgColor : '#808080'}}
      // sheet1.getRow(8).values = ['S.No.', 'PART CODE', 'DESCRIPTION', 'SPECIFICATION', 'TYPE', 'QTY/AHU', 'UOM', 'TOTAL QTY'];

      sheet1.getCell('A9').value = 'AHU CASING';
      sheet1.columns = [
        { key: 'sno', width: 15 },
        { key: 'PART_CODE', width: 15 },
        { key: 'DESCRIPTION', width: 15 },
        { key: 'SPECIFICATION', width: 15 },
        { key: 'TYPE', width: 15 },
        { key: 'QTY', width: 15 },
        { key: 'UOM', width: 15 },
        { key: 'TOTAL_QTY', width: 15 }]

      ahuCasingData.forEach(function (item, index) {
        sheet1.addRow({
          sno: index + 1,
          PART_CODE: item.part_code,
          DESCRIPTION: item.description,
          SPECIFICATION: item.specification,
          TYPE: item.type,
          QTY: item.qty,
          UOM: item.uom,
          TOTAL_QTY: item.totalQty,
        })
      })
      var lastRow = sheet1.lastRow;
      // console.log(lastRow._number+1);
      var getRowInsert = sheet1.getRow(++(lastRow.number));
      console.log(getRowInsert);
      console.log('working sheet');
      sheet1.addRow('BLANK OFF & OTHERS');
      // sheet1.getCell('B10:H24').value = ahuCasingData
      workbook.xlsx.writeFile('./data/calc/new_bom.xlsx');
      // sending file
      let file = path.join(`${__dirname}/data/calc/new_bom.xlsx`);
      res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.download(file, 'New Bom.xlsx');
    });

  });

})

app.get('/download', function (req, res) {
  var file = path.join(`${__dirname}/data/calc/new_bom.xlsx`);
  res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.download(file, 'New Bom.xlsx');
})



app.get('/fetchCasing', function (req, res) {
  var wb = XLSX.readFile('formula.xlsx', { bookDeps: true });
  var ws = wb.Sheets['test'];
  var stream = XLSX.stream.to_json(ws);
  var data = XLSX.utils.sheet_to_json(ws);
  res.send({ s: [], r: data })
});





// var passData = XLSX.utils.

// console.log(ws.K2.v);




// var data = XLSX.utils.sheet_to_json(ws);
// console.log(data);
// function add_cell_to_sheet(worksheet, address, value) {
// 	/* cell object */
// 	var cell = {t:'?', v:value};

// 	/* assign type */
// 	if(typeof value == "string") cell.t = 's'; // string
// 	else if(typeof value == "number") cell.t = 'n'; // number
// 	else if(value === true || value === false) cell.t = 'b'; // boolean
// 	else if(value instanceof Date) cell.t = 'd';
// 	else throw new Error("cannot store value");

// 	/* add to worksheet, overwriting a cell if it exists */
// 	worksheet[address] = cell;

// 	/* find the cell range */
// 	var range = XLSX.utils.decode_range(worksheet['!ref']);
// 	var addr = XLSX.utils.decode_cell(address);

// 	/* extend the range to include the new cell */
// 	if(range.s.c > addr.c) range.s.c = addr.c;
// 	if(range.s.r > addr.r) range.s.r = addr.r;
// 	if(range.e.c < addr.c) range.e.c = addr.c;
// 	if(range.e.r < addr.r) range.e.r = addr.r;

// 	/* update range */
// 	worksheet['!ref'] = XLSX.utils.encode_range(range);
// }

// add_cell_to_sheet(ws, "F6", 12345);

// XLSX.writeFile('sheetjs-new.xlsx', wb);


// const jsonParser = bodyParser.json();

// app.get('/itemMasterData',jsonParser, function(req, res){
//   var workbook = XLSX.readFile('./data/item_master.xlsx');
// var sheet_name_list = workbook.SheetNames;
// res.send(XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]))
// // console.log(XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[1]]))
// // sheet_name_list.forEach(function(y) {
// //     var worksheet = workbook.Sheets[y];
// //     var headers = {};
// //     var data = [];
// //     for(z in worksheet) {
// //         if(z[0] === '!') continue;
// //         //parse out the column, row, and value
// //         var tt = 0;
// //         for (var i = 0; i < z.length; i++) {
// //             if (!isNaN(z[i])) {
// //                 tt = i;
// //                 break;
// //             }
// //         };
// //         var col = z.substring(0,tt);
// //         var row = parseInt(z.substring(tt));
// //         var value = worksheet[z].v;

// //         //store header names
// //         if(row == 1 && value) {
// //             headers[col] = value;
// //             continue;
// //         }

// //         if(!data[row]) data[row]={};
// //         data[row][headers[col]] = value;
// //     }
// //     //drop those first two rows which are empty
// //     data.shift();
// //     data.shift();
// //     res.send(data);
// // });
// });

app.use('/', (req, res) => { res.send('Welcome to AHU-BOM Data Server') });
app.listen(3200, () => console.log('Server is running at PORT no. :' + 3200));
