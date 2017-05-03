const electron = require('electron')
// Module to control application life.
const app = electron.app
// Module to create native browser window.
const BrowserWindow = electron.BrowserWindow

const path = require('path')
const url = require('url')

// Keep a global reference of the window object, if you don't, the window will
// be closed automatically when the JavaScript object is garbage collected.
let mainWindow

function createWindow () {
  // Create the browser window.
  mainWindow = new BrowserWindow({width: 800, height: 600})

  // and load the index.html of the app.
  mainWindow.loadURL(url.format({
    pathname: path.join(__dirname, 'index.html'),
    protocol: 'file:',
    slashes: true
  }))

  // Open the DevTools.
  // mainWindow.webContents.openDevTools()

  // Emitted when the window is closed.
  mainWindow.on('closed', function () {
    // Dereference the window object, usually you would store windows
    // in an array if your app supports multi windows, this is the time
    // when you should delete the corresponding element.
    mainWindow = null
  })
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.on('ready', createWindow)

// Quit when all windows are closed.
app.on('window-all-closed', function () {
  // On OS X it is common for applications and their menu bar
  // to stay active until the user quits explicitly with Cmd + Q
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', function () {
  // On OS X it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (mainWindow === null) {
    createWindow()
  }
})

// In this file you can include the rest of your app's specific main process
// code. You can also put them in separate files and require them here.
// 开始业务代码
const ipc = require('electron').ipcMain
const dialog = require('electron').dialog
const xlsx = require('xlsx')
const fs = require('fs')

function to_json(workbook) {
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var roa = xlsx.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
		if(roa.length > 0){
			result[sheetName] = roa;
		}
	});
	return result;
}


ipc.on('open-file-dialog', function (event) {
    dialog.showOpenDialog({
        properties: ['openFile', 'openDirectory']
    }, function (path) {
        if(typeof path === 'undefined'){
            return
        }
        fs.readdir(path+'/', function(err, files){
            if(files.length === 0){
                event.sender.send('result', '')
                return
            }
            files = files.filter(function(val){
                // 匹配文件名称格式是否正确
                return /^\d+_\d+\.xls$/.test(val)
            })
            if(files.length === 0){
                event.sender.send('result', '')
                return
            }
            var partnerSale = {}
            files.forEach(function(file){
                // event.sender.send('result', {type:'checkfile', data:file})

                var workbook = xlsx.readFile(path+'/'+file)
                var data = to_json(workbook)
                if(typeof data['所有IC卡交易明细'] === undefined){
                    return
                }
                data = data['所有IC卡交易明细']

                data.forEach(function(val){
                    if( val['业务类型'] === '加油' && typeof val['金额(分值)'] !== 'undefined'){
                        if(typeof partnerSale[val['地点']] === 'undefined'){
                            partnerSale[val['地点']] = {
                                totalSale: 0,
                                bonus: 0
                            }
                        }
                        partnerSale[val['地点']]['totalSale'] += parseFloat(val['金额(分值)'], 10)
                        partnerSale[val['地点']]['bonus'] += parseFloat(val['奖励分值'], 10)
                    }
                })
            })
            if(partnerSale === {}){
                return
            }

            var _headers = ['站点名称', '消费金额', '奖励分']
            var _data = []
            for(var i in partnerSale){
                _data.push({
                    '站点名称': i,
                    '消费金额': "_" + Math.round(partnerSale[i]['totalSale']),
                    '奖励分': "_" + Math.round(partnerSale[i]['bonus'])
                })
            }
            var headers = _headers.map((v, i) => Object.assign({}, {v: v, position: String.fromCharCode(65+i) + 1 })).reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
            var data = _data.map((v, i) => _headers.map((k, j) => Object.assign({}, { v: v[k], position: String.fromCharCode(65+j) + (i+2) }))).reduce((prev, next) => prev.concat(next)).reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
            var output = Object.assign({}, headers, data);
            var outputPos = Object.keys(output);
            var ref = outputPos[0] + ':' + outputPos[outputPos.length - 1];
            // var finaldata = [];
            newWorkbook = {
                SheetNames: ['haha'],
                Sheets:{
                    'haha': Object.assign({}, output, { '!ref': ref })
                }
            }
            xlsx.writeFile(newWorkbook, path + '/final.xlsx')

            event.sender.send('result', '请查看生成excel文件'+path + '/final.xlsx')

        })
    })
})
