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
        var desired_cell = workbook.Sheets[sheetName]['E1']
        var desired_value = (desired_cell ? desired_cell.v : undefined)
        if( desired_value.trim() === '中国石化加油IC卡台帐对帐单'){
            // 删除前面5行
            var roa = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], {range: 5});
        } else {
            var roa = xlsx.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
        }
		if(roa.length > 0){
			result[sheetName] = roa;
		}
	});
	return result;
}

ipc.on('open_file_dialog', function (event) {
    dialog.showOpenDialog({
        properties: ['openFile','openDirectory']
    }, function (choosePath) {
        event.sender.send('accept_dir_path', choosePath)
        if(typeof choosePath !== 'undefined' ){
            // 新建个窗口执行后台任务
            var win = new BrowserWindow({
                width: 400, height: 225, show: false
            })
            win.loadURL(url.format({
                pathname: path.join(__dirname, 'invisible.html'),
                protocol: 'file:',
                slashes: true
            }))
            // win.webContents.openDevTools()
            win.webContents.on('did-finish-load', function () {
                win.webContents.send('bg_accept_path', choosePath)
            })
        }
    })
})


function getAllFiles(root){
  var res = [] , files = fs.readdirSync(root);
  files.forEach(function(file){
    var pathname = root+'/'+file
    , stat = fs.lstatSync(pathname);

    if (!stat.isDirectory()){
        // 查找过滤积分文件和财务统计的excel
        if( /^.*\.xls$/.test(pathname) &&  pathname.indexOf('财务统计')=== -1 && pathname.indexOf('jf')=== -1 ){
            res.push(pathname);
        }
    } else {
        // 检查目录后缀
        if(pathname.indexOf('jf')=== -1 && pathname.indexOf('统计')=== -1){
            res = res.concat(getAllFiles(pathname));
        }
    }
  });
  return res
}

ipc.on('choose_dir', function (event, arg) {
    var choosePath = arg[0]
    var files = []
    files = getAllFiles(choosePath)
    if(files.length === 0){
        mainWindow.webContents.send('read_path', '没有找到任何文件')
        event.sender.send('read_path', '没有找到任何文件')
        return
    }

    var partnerSale = {}
    files.forEach(function(file){
        mainWindow.webContents.send('read_path_one_by_one', file)
        var workbook = xlsx.readFile(file, {sheetRows:5100})
        var data = to_json(workbook)
        workbook = null
        if(typeof data['所有IC卡交易明细'] === undefined){
            return
        }
        data = data['所有IC卡交易明细']

        data.forEach(function(val){
            if( val['业务类型'] === '加油' && typeof val['金额(分值)'] !== 'undefined'){
                if(typeof partnerSale[val['地点']] === 'undefined'){
                    partnerSale[val['地点']] = {
                        totalSale: 0,
                        bonus: 0,
                        num: 0
                    }
                }
                partnerSale[val['地点']]['totalSale'] += parseFloat(val['金额(分值)'], 10)
                partnerSale[val['地点']]['bonus'] += parseFloat(val['奖励分值'], 10)
                partnerSale[val['地点']]['num'] += parseFloat(val['数量'], 10)
            }
        })


    })



    if( Object.keys(partnerSale).length === 0){
        mainWindow.webContents.send('read_path', '并没有统计到任何符合规则的信息')
        event.sender.send('read_path', '并没有统计到任何符合规则的信息')
        return
    }

    var _headers = ['站点名称', '消费金额', '奖励分', '加油量']
    var _data = []
    for(var i in partnerSale){
        _data.push({
            '站点名称': i,
            '消费金额': "_" + Math.round(partnerSale[i]['totalSale']),
            '奖励分': "_" + Math.round(partnerSale[i]['bonus']),
            '加油量': "_" + Math.round(partnerSale[i]['num'])
        })
    }
    var headers = _headers.map((v, i) => Object.assign({}, {v: v, position: String.fromCharCode(65+i) + 1 })).reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
    var data = _data.map((v, i) => _headers.map((k, j) => Object.assign({}, { v: v[k], position: String.fromCharCode(65+j) + (i+2) }))).reduce((prev, next) => prev.concat(next)).reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
    var output = Object.assign({}, headers, data);
    var outputPos = Object.keys(output);
    var ref = outputPos[0] + ':' + outputPos[outputPos.length - 1];
    // var finaldata = [];
    newWorkbook = {
        SheetNames: ['财务统计'],
        Sheets:{
            '财务统计': Object.assign({}, output, { '!ref': ref })
        }
    }
    xlsx.writeFile(newWorkbook, choosePath + '/财务统计.xlsx')
    mainWindow.webContents.send('read_path', '已经成功生成excel文件!!!'+choosePath + '/财务统计.xlsx')
    event.sender.send('read_path', '已经成功生成excel文件!!!'+choosePath + '/财务统计.xlsx')
})
