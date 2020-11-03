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
const moment = require('moment');
// var mime = require('mime');
// const Blob = require('cross-blob');
// const { request } = require('http');
// request = require('request');

const FormulaParser = require('hot-formula-parser').Parser;
const parser = new FormulaParser();

// console.log(moment().format("DD-MMM-YYYY"));

const Excel = require('exceljs');
// const { style } = require('@angular/animations');
const workbook = new Excel.Workbook();
const calcworkbook = new Excel.Workbook();

const jsonParser = bodyParser.json();

app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());
app.use(cors());
// xlsxFile('./data/item_master.xlsx').then((rows) => {
// });



app.post('/casingCalc', function (req, res) {
  let calcdata = [];
  let calcDataCoil = [];
  var workbookItem = XLSX.readFile('./data/item_master.xlsx');
  var sheet_name_list1 = workbookItem.SheetNames;
  var imData = XLSX.utils.sheet_to_json(workbookItem.Sheets[sheet_name_list1[0]])
  // res.send(imData);
  // console.log(req.body);
  let wb = XLSX.readFile('./data/calc/casing.xlsx');
  let ws = wb.Sheets['Casing'];
  let wsCoil = wb.Sheets['Coil'];
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

  

  // function getCellResult2(coilSheet, cellLabel) {
  //   if (coilSheet.getCell(cellLabel).formula) {
  //     return parser.parse(coilSheet.getCell(cellLabel).formula).result;
  //   } else {
  //     return coilSheet.getCell(cellCoord.label).value;
  //   }
  // }




  workbook.xlsx.readFile('./data/calc/bom_temp.xlsx').then(() => {

    calcworkbook.xlsx.readFile('./data/calc/casing.xlsx').then(() => {
      var worksheet = calcworkbook.getWorksheet(1);
      var coilSheet = calcworkbook.getWorksheet(2);
      console.log(coilSheet.getCell('A2').value, 'kkkkk');
      // console.log(req.body);
      worksheet.getCell('A2').value = req.body.unitForm.innerSheet.Description;
      worksheet.getCell('B2').value = req.body.unitForm.outerSheet.Description;
      coilSheet.getCell('A2').value = req.body.coilForm.coils[0].finHeight;
      coilSheet.getCell('B2').value = req.body.coilForm.coils[0].finLength;
      coilSheet.getCell('C2').value = req.body.coilForm.coils[0].rowDeep;
      coilSheet.getCell('E2').value = req.body.coilForm.coils[0].fpi;
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
        // if (coilSheet.getCell(cellCoord.label).formula) {
        //   done(parser.parse(coilSheet.getCell(cellCoord.label).formula).result);
        // } else {
        //   done(coilSheet.getCell(cellCoord.label).value);
        // }
      });

      parser.on('callRangeValue', function (startCellCoord, endCellCoord, done) {
        var fragment = [];
        
        var fragmentCoil = [];

        for (var row = startCellCoord.row.index; row <= endCellCoord.row.index; row++) {
          var colFragment = [];
          
          // var colFragmentCoil = [];

          for (var col = startCellCoord.column.index; col <= endCellCoord.column.index; col++) {
            colFragment.push(worksheet.getRow(row + 1).getCell(col + 1).value);
            // colFragmentCoil.push(coilSheet.getRow(row + 1).getCell(col + 1).value);
          }

          fragment.push(colFragment);
          // fragmentCoil.push(colFragmentCoil);
          // console.log(fragment, 'Testing parsing');
        }

        if (fragment) {
          done(fragment);
        }
        if (fragmentCoil) {
          done(fragmentCoil);
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
      calcdata[0].area = (area).toFixed(2);
      calcdata[0].inner_sheet_weight = (isw).toFixed(2);
      calcdata[0].outer_sheet_weight = osw.toFixed(2);
      calcdata[0].corner_profile = O.toFixed(2);
      calcdata[0].omega_profile = P.toFixed(2);
      calcdata[0].polyol = Q.toFixed(2);
      calcdata[0].isol = R.toFixed(2);

      
      // var tube_N = getCellResult(coilSheet,'C2')
      // console.log('No. of Tubes',tube_N);

      // var tube_N = getCellResult(worksheet, 'C2');

      calcDataCoil = XLSX.utils.sheet_to_json(wsCoil);

      // console.log(tube_N);

      // console.log(calcdata);
      // res.send(calcdata);
      // console.log(workbook);

      
      //Casing Iteration ---Starts----
      var profiles = imData.filter(i => (i.Item_Group === 'ALUMINIUM PROFILES'));
      
      var plasticParts = imData.filter(i => (i.Item_Group === 'PLASTIC PARTS'));

      
      var polyol = imData.filter(i => (i.Item_Group === 'POLYOL'));
      var isol = imData.filter(i => (i.Item_Group === 'ISO-CYNATE'));
      var gasket = imData.filter(i => (i.Item_Group === 'GASKET & INSULATION'));
      var fastners = imData.filter(i => (i.Item_Group === 'FASTENERS'));
      var fins = imData.filter(i => (i.Sub_Group === 'Fin'));
      var cu_tubes = imData.filter(i => (i.Sub_Group === 'COPPER TUBE'));
      var U_bend_data = imData.filter(i => (i.Sub_Group === 'U_Bend'));
      var coil_casing_data = imData.filter(i => (i.Sub_Group === 'Coil Casing'));
      var coil_header_data = imData.filter(i => (i.Sub_Group === 'Coil Header'));
      var drainPlug_data = imData.filter(i => (i.Sub_Group === 'COIL DRAIN'));
      var purge = imData.filter(i => (i.Sub_Group === 'COIL PURGE'));

      console.log(U_bend_data, 'UUUU');

      let ahuCasingData = [];
      let supplyFanData = [];
      let coilData = [];
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
      var ProfileData = profiles.filter(i => (i.Description === req.body.unitForm.profileType));
      // console.log(ProfileData);
      var cornerProfile = {
        part_code: ProfileData[0].Code,
        description: 'Corner Profile', specification: ProfileData[0].Name, type: '',
        qty: calcdata[0].corner_profile, uom: ProfileData[0].Unit, totalQty: calcdata[0].corner_profile * req.body.ahuQty
      }
      var omegaProfile = {
        part_code: ProfileData[1].Code,
        description: 'Omega Profile', specification: ProfileData[1].Name, type: '',
        qty: calcdata[0].omega_profile, uom: ProfileData[1].Unit, totalQty: calcdata[0].omega_profile * req.body.ahuQty
      }
      var JoinerData = plasticParts.filter(i => (i.Description == req.body.unitForm.profileType));
      // console.log(JoinerData, 'fdsf');
      var cornerJoiner = {
        part_code: JoinerData[0].Code,
        description: 'Corner Joiner', specification: JoinerData[0].Name, type: '',
        qty: 8, uom: JoinerData[0].Unit, totalQty: 8 * req.body.ahuQty
      }
      var omegaJoiner = {
        part_code: JoinerData[1].Code,
        description: 'Omega Joiner', specification: JoinerData[1].Name, type: '',
        qty: 44, uom: JoinerData[1].Unit, totalQty: 44 * req.body.ahuQty
      }
      
      ahuCasingData.push(innerSkin, outerSkin, cornerProfile,omegaProfile,cornerJoiner, omegaJoiner);
      
      if(req.body.unitForm.insulation == 'PUF'){
        var insulation_polyol = {
          part_code: polyol[0].Code,
          description: 'Polyol', specification: polyol[0].Name, type: '',
          qty: (calcdata[0].polyol), uom: polyol[0].Unit, totalQty: (calcdata[0].polyol) * req.body.ahuQty
        }
        var insulation_isol = {
          part_code: isol[0].Code,
          description: 'Isol', specification: isol[0].Name, type: '',
          qty: (calcdata[0].isol), uom: isol[0].Unit, totalQty: (calcdata[0].isol) * req.body.ahuQty
        }
        ahuCasingData.push(insulation_polyol, insulation_isol);
      } else {
        var rockwool = {
          part_code: '',
          description: 'RockWool', specification: '', type: '',
          qty: '', uom: '', totalQty: ''
        }
        ahuCasingData.push(rockwool)
      }

      var gasket_panel = gasket.filter(i => (i.Description === 'Panel'));
      var fastners_panel = fastners.filter(i => (i.Description === 'Panel'));
      const gasketQty = 8

      var gasketPanel_entry ={
        part_code: gasket_panel[0].Code,
        description: 'EPDM Panel Gasket', specification: gasket_panel[0].Name, type: '',
        qty: gasketQty, uom: gasket_panel[0].Unit, totalQty: gasketQty * req.body.ahuQty
      }
      var fastnersPanel_entry1 ={
        part_code: fastners_panel[0].Code,
        description: 'Panel fixing Screws', specification: fastners_panel[0].Name, type: '',
        qty: (gasketQty*10/0.3), uom: fastners_panel[0].Unit, totalQty: (gasketQty*10/0.3) * req.body.ahuQty
      }
      var fastnersPanel_entry2 ={
        part_code: fastners_panel[0].Code,
        description: 'Blank off, Omega , drain tray fixing Screws', specification: fastners_panel[0].Name, type: '',
        qty: 100, uom: fastners_panel[0].Unit, totalQty: 100 * req.body.ahuQty
      }

      ahuCasingData.push(gasketPanel_entry, fastnersPanel_entry1, fastnersPanel_entry2)

      //Casing Iteration ---Ends----

      //Fan Iteration ----Starts----

      var supplyFanModel = {
        part_code: req.body.fanForm.supplyFan.fanModel.Code,
        description: 'Blower', specification: req.body.fanForm.supplyFan.fanModel.Name, type: '',
        qty: req.body.fanForm.supplyFan.fanQty, uom: req.body.fanForm.supplyFan.fanModel.Unit, totalQty: req.body.fanForm.supplyFan.fanQty * req.body.ahuQty
      }
      var supplyMotorModel = {
        part_code: req.body.fanForm.supplyFan.motorModel.Code,
        description: 'Motor', specification: req.body.fanForm.supplyFan.motorModel.Name, type: '',
        qty: req.body.fanForm.supplyFan.motorQty, uom: req.body.fanForm.supplyFan.motorModel.Unit, totalQty: req.body.fanForm.supplyFan.motorQty * req.body.ahuQty
      }
      var supplyAntiVibrant = {
        part_code: req.body.fanForm.supplyFan.antiVibrant.Code,
        description: 'Anti Vibrant', specification: req.body.fanForm.supplyFan.antiVibrant.Name, type: '',
        qty: 4, uom: req.body.fanForm.supplyFan.antiVibrant.Unit, totalQty: 4 * req.body.ahuQty
      }
      var supplyFanPulley = {
        part_code: req.body.fanForm.supplyFan.fanPulley.Code,
        description: 'Fan Pulley', specification: req.body.fanForm.supplyFan.fanPulley.Name, type: '',
        qty: 1, uom: req.body.fanForm.supplyFan.fanPulley.Unit, totalQty: 1 * req.body.ahuQty
      }
      var supplyMotorPulley = {
        part_code: req.body.fanForm.supplyFan.motorPulley.Code,
        description: 'Motor Pulley', specification: req.body.fanForm.supplyFan.motorPulley.Name, type: '',
        qty: 1, uom: req.body.fanForm.supplyFan.motorPulley.Unit, totalQty: 1 * req.body.ahuQty
      }
      var supplyBelt = {
        part_code: req.body.fanForm.supplyFan.belt.Code,
        description: 'Belt', specification: req.body.fanForm.supplyFan.belt.Name, type: '',
        qty: 2, uom: req.body.fanForm.supplyFan.belt.Unit, totalQty: 2 * req.body.ahuQty
      }

      supplyFanData.push('',supplyFanModel,supplyMotorModel,supplyAntiVibrant,supplyBelt, supplyFanPulley, supplyMotorPulley);

      //Fan Iteration ----Ends----

      //Coil Iteration ----Starts----
      // let fin;
      // if(req.body.coilForm.coils[0].coilType.Name === 'WATER COOLING COIL' && req.body.coil )
      var cd = calcDataCoil[0]
      var tube_N = Math.round(cd.fin_height*cd.tubes_f1)
      var fin_qty = ((cd.fin_height/cd.meter_conv)*(cd.fin_length/cd.meter_conv)*cd.rd*cd.fin_qty_f1*cd.fin_qty_f2).toFixed(2)
      var tube_qty = ((cd.fin_height/cd.meter_conv)*(cd.fin_length/cd.meter_conv)*cd.rd*cd.fin_qty_f1*cd.fin_qty_f2).toFixed(2)
      var casing_qty = (((cd.fin_height+cd.fin_length)*cd.casing_qty_f1/cd.meter_conv)*((cd.rd+cd.casing_qty_f2)*cd.casing_qty_f3/cd.meter_conv)*cd.casing_qty_f4*cd.casing_qty_f5*cd.casing_qty_f6).toFixed(2)
      var header_qty = Math.round((cd.fin_height*cd.header_qty_f1+cd.header_qty_f2)*cd.header_qty_f3/cd.meter_conv)
      var drainpan_sheet_qty = Math.round((cd.drainpan_qty_f1*cd.fin_length/cd.meter_conv*cd.drainpan_qty_f2*cd.drainpan_qty_f3))
      var U_bend_qty = Math.round(cd.fin_height/cd.tubes_f1*(cd.rd*cd.U_bend_f1-cd.U_bend_f2)/cd.U_bend_f3)
      var copper_stub_qty = (tube_N*cd.copper_st_stub_f1*cd.copper_st_stub_f2).toFixed(2)

      var fin = {
        part_code: req.body.coilForm.coils[0].finMaterial.Code,
        description: 'Fin', specification: req.body.coilForm.coils[0].finMaterial.Name, type: '',
        qty: fin_qty, 
        uom: req.body.coilForm.coils[0].finMaterial.Unit, totalQty: calcDataCoil[0].fin_qty * req.body.ahuQty
      }
      var tube = {
        part_code: req.body.coilForm.coils[0].tubeDia.Code,
        description: 'Copper Tube', specification: req.body.coilForm.coils[0].tubeDia.Name, type: '',
        qty: tube_qty,
        uom: req.body.coilForm.coils[0].tubeDia.Unit, totalQty: calcDataCoil[0].tube_qty * req.body.ahuQty
      }
      var casingCoil = {
        part_code: coil_casing_data[0].Code,
        description: 'End Plate & Side Plates', specification: coil_casing_data[0].Name, type: '',
        qty: casing_qty,
        uom: coil_casing_data[0].Unit, totalQty: casing_qty * req.body.ahuQty
      }
      var headerCoil = {
        part_code: req.body.coilForm.coils[0].headerMaterial.Code,
        description: 'Header', specification: req.body.coilForm.coils[0].headerMaterial.Name, type: '',
        qty: header_qty,
        uom: req.body.coilForm.coils[0].headerMaterial.Unit, totalQty: header_qty * req.body.ahuQty
      }
      // console.log(U_bend_data);
      var U_bend = {
        part_code: U_bend_data[0].Code,
        description: '"U" Bend', specification: U_bend_data[0].Name, type: '',
        qty: U_bend_qty,
        uom: U_bend_data[0].Unit, totalQty: U_bend_qty * req.body.ahuQty
      }
      var copper_stub = {
        part_code: cu_tubes[1].Code,
        description: 'Copper Straight Stub', specification: cu_tubes[1].Name, type: '',
        qty: copper_stub_qty,
        uom: cu_tubes[1].Unit, totalQty: copper_stub_qty * req.body.ahuQty
      }
      var drainPlug = {
        part_code: drainPlug_data[0].Code,
        description: 'Drain Plug', specification: drainPlug_data[0].Name, type: '',
        qty: 1,
        uom: drainPlug_data[0].Unit, totalQty: 1 * req.body.ahuQty
      }
      var coilPurge = {
        part_code: purge[0].Code,
        description: 'Drain Plug', specification: purge[0].Name, type: '',
        qty: req.body.coilForm.coils[0].circuit,
        uom: purge[0].Unit, totalQty: req.body.coilForm.coils[0].circuit * req.body.ahuQty
      }


      coilData.push(fin, tube, casingCoil, headerCoil,  U_bend, copper_stub, drainPlug, coilPurge)
      //Coil Iteration ----Ends----

      var sheet1 = workbook.getWorksheet(1);
      sheet1.columns
      // console.log(sheet1.getCell('A1').value);
      sheet1.getCell('A1').value = 'Bill Of Materials';
      sheet1.getCell(3, 3).value = req.body.project;
      sheet1.getCell(4, 3).value = req.body.ahuType;
      sheet1.getCell(5, 3).value = req.body.unitForm.airVolume;
      sheet1.getCell('C6').value = req.body.ahuQty;
      // sheet1.getCell('C6').formula
      sheet1.getCell('C7').value = req.body.ahuModel;
      sheet1.getCell('D6').value = 'DATE                  :    ' + (moment().format("DD-MMM-YYYY"));
      sheet1.mergeCells('A'+9+':H9');
      sheet1.getCell('A9:H9').style = {border: {bottom: {style: 'thin'}}};
      sheet1.getCell('A9').alignment = { horizontal: 'center', vertical: 'middle' };
      // sheet1.getRow(9).style = {border: {bottom: {style: 'thin'}}};
      // sheet1.getCell('A9').style = {fill: {bgColor : '#808080'}}
      // sheet1.getRow(8).values = ['S.No.', 'PART CODE', 'DESCRIPTION', 'SPECIFICATION', 'TYPE', 'QTY/AHU', 'UOM', 'TOTAL QTY'];

      sheet1.getCell('A9').value = 'AHU CASING';;
      // sheet1.mergeCells(fanHeadCell);
      sheet1.columns = [
        { key: 'sno', width: 15, style:{border: {bottom: {style : 'thin'}}} },
        { key: 'PART_CODE', width: 15 },
        { key: 'DESCRIPTION', width: 45 },
        { key: 'SPECIFICATION', width: 70 },
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
      });
      
      supplyFanData.forEach(function (item, index) {
        sheet1.addRow({
          sno: (index + ahuCasingData.length) + 1,
          PART_CODE: item.part_code,
          DESCRIPTION: item.description,
          SPECIFICATION: item.specification,
          TYPE: item.type,
          QTY: item.qty,
          UOM: item.uom,
          TOTAL_QTY: item.totalQty,
        })
      });
      
      var fanHeadPosition = 9 + ahuCasingData.length + 1
      var fanHeadCell = 'A'+fanHeadPosition+':H'+fanHeadPosition
      console.log(fanHeadCell, typeof(fanHeadCell));

      console.log(coilData, 'coilD');

      sheet1.getCell('A'+fanHeadPosition).value = 'Supply Fan';
      sheet1.mergeCells(fanHeadCell);
      sheet1.getCell('A'+fanHeadPosition).alignment = { horizontal: 'center', vertical: 'middle' }

      var wCoilHeadPosition = 9 + ahuCasingData.length +supplyFanData.length+ 1
      var wCoilHeadCell = 'A'+wCoilHeadPosition+':H'+wCoilHeadPosition
      console.log(wCoilHeadCell, typeof(wCoilHeadCell));

      sheet1.getCell('A'+wCoilHeadPosition).value = req.body.coilForm.coils.length+' - '+ req.body.coilForm.coils[0].coilType.Code+' - '+ req.body.coilForm.coils[0].coilType.Name+' ( Code : '+ req.body.coilForm.coils[0].finHeight+'H X '+req.body.coilForm.coils[0].finHeight+'L X '+req.body.coilForm.coils[0].rowDeep+' R X '+ req.body.coilForm.coils[0].fpi+' FPI X '+ req.body.coilForm.coils[0].tubeDia.Description+' DIA X '+ req.body.coilForm.coils[0].circuit+' C )';
      sheet1.mergeCells(wCoilHeadCell);
      sheet1.getCell('A'+wCoilHeadPosition).alignment = { horizontal: 'center', vertical: 'middle' }
      coilData.forEach(function (item, index) {
        console.log('coilIterate');
        sheet1.addRow({
          sno: (index + ahuCasingData.length)+supplyFanData.length + 1,
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
      // console.log(getRowInsert);
      // console.log('working sheet');
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
  console.log('Hi');
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

app.get('/itemMasterData',jsonParser, function(req, res){
  var workbookItem = XLSX.readFile('./data/item_master.xlsx');
var sheet_name_list1 = workbookItem.SheetNames;
var imData = XLSX.utils.sheet_to_json(workbookItem.Sheets[sheet_name_list1[0]])
var profiles = imData.filter(i => (i.Item_Group === 'ALUMINIUM PROFILES'));
let profile_desc = []
profiles.forEach(i => {
  profile_desc.push(i.Description);
})
      // console.log(profiles);
res.send(imData);
// console.log(XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[1]]))
// sheet_name_list.forEach(function(y) {
//     var worksheet = workbook.Sheets[y];
//     var headers = {};
//     var data = [];
//     for(z in worksheet) {
//         if(z[0] === '!') continue;
//         //parse out the column, row, and value
//         var tt = 0;
//         for (var i = 0; i < z.length; i++) {
//             if (!isNaN(z[i])) {
//                 tt = i;
//                 break;
//             }
//         };
//         var col = z.substring(0,tt);
//         var row = parseInt(z.substring(tt));
//         var value = worksheet[z].v;

//         //store header names
//         if(row == 1 && value) {
//             headers[col] = value;
//             continue;
//         }

//         if(!data[row]) data[row]={};
//         data[row][headers[col]] = value;
//     }
//     //drop those first two rows which are empty
//     data.shift();
//     data.shift();
//     res.send(data);
// });
});

app.use('/', (req, res) => { res.send('Welcome to AHU-BOM Data Server') });
app.listen(3200, () => console.log('Server is running at PORT no. :' + 3200));
