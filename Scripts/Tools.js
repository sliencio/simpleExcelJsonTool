var path = require('path');
var fs = require('fs');
var Excel = require('exceljs');
var xlsx = require('node-xlsx');
var config = require('../config.json');
var rootPath = path.join(__dirname, "../");

/**
 * 获取路径
 * @param  {...any} args 
 */
function GetPath(...args) {
    var pathStr = "";
    for (let i = 0, len = args.length; i < len; i++) {
        pathStr += args[i];
    }
    return path.join(rootPath, pathStr);
}


/**
 * @param 
 *      [{ header: '第一行第一个字段', key: 'Id', width: 10 },
        { header: '第一行第二个字段', key: 'Answers', width: 100 },
        { header: '最大', key: 'Max', width: 10 },
        { header: '最小', key: 'Min', width: 10 },
        { header: '和上一级的差', key: 'Subtract' }]
* header 
 * @param {固定行数据} fixRowData 
 * @param {需要填充的数据} dataArray 
 * @param {导出路径} exportPath 
 */
function ExportExcel(fileName) {
    var jsonPath = GetPath(config.jsonDirectoryPath + fileName + ".json") ;
    //同步读取
    var rawData = fs.readFileSync(jsonPath);
    var jsonData = JSON.parse(rawData.toString());
    for (var sheetName in jsonData) {
        var workbook = new Excel.stream.xlsx.WorkbookWriter({
            filename: GetPath(config.exportExcelPath, sheetName + ".xlsx")
        });
        //添加页签
        var worksheet = workbook.addWorksheet(sheetName);
        //设置行
        console.log("开始导出 " + sheetName+"==> Excel");
        // 开始添加数据
        for (let rowDataKey in jsonData[sheetName]) {
            var rowData = jsonData[sheetName][rowDataKey];
            //如果是对象
            if (!Array.isArray(rowData)) {
                var dataArray = new Array();
                for (var key in rowData) {
                    dataArray.push(rowData[key]);
                }
                worksheet.addRow(dataArray).commit();
            } else {
                for (var i = 0; i < rowData.length; i++){
                    worksheet.addRow(rowData[i]).commit();
                }
            }
            
        }
        workbook.commit();
        console.log("成功导出" + sheetName +" Excel ==  ╮(￣▽￣)╭");
    }
}

/**
 * 导出json文件
 * @param {json名称} jsonName 
 * @param {导出数据} data 
 */
function ExportJson(jsonName, dataString) {
    //导出使用文件
    fs.writeFile(config.exportjsonPath + jsonName + ".json", dataString, function (err) {
        if (err) {
            console.log('导出 ' + jsonName + ' 失败');
        }
        else {
            console.log('导出 ' + jsonName + ' 成功 ╮(￣▽￣)╭ ');
        }
    });
}

/**
 * 
 * @param {路径} excelPath
 */
function ParseExcel(excelPath) {
    var rawData = xlsx.parse(excelPath);
    var retData = {};
    var excelName = path.basename(excelPath, '.xlsx');

    rawData.forEach(sheetData => {
        var name = sheetData.name;
        retData[name] = {};
        var fieldRowData = sheetData.data[config.fieldIndex - 1];
        var tempIndex = 1;
        for (var i = config.beginReadRow - 1; i < sheetData.data.length; i++) {
            var tempData = {};
            var rowDataArray = sheetData.data[i];

            for (var j = 0; j < fieldRowData.length; j++) {
                tempData[fieldRowData[j]] = rowDataArray[j];
            }
            if (config.PrimaryIndex <= 0) {
                retData[name][tempIndex.toString()] = tempData;
                tempIndex++;
            } else {
                retData[name][rowDataArray[config.PrimaryIndex - 1]] = tempData;
            }
        }
    });
    ExportJson(excelName, JSON.stringify(retData, null, config.formatJson ? 4 : 0));
    // console.log(JSON.stringify(retData));
}

module.exports = {
    GetPath: GetPath,
    ExportExcel: ExportExcel,
    ParseExcel: ParseExcel
}