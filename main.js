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
	let result = {};
	workbook.SheetNames.forEach(function(sheetName) {
        if(sheetName !== '所有IC卡交易明细'){
            return
        }
        let roa = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], {header: 'A'});

        roa.forEach(function(val){
            // 有两种格式的excel,吗的,现在不记得是哪两种格式的excel了!!!!!!!
            let yewuleixing = 'C'
            let jianglifenzhi = 'H'
            let jine = 'G'
            let zhandian = 'K'
            let kahao = 'A'
            let shijian = 'B'
            let pinzhong = 'D'
            if(val['D'] === '加油'){
                yewuleixing = 'D'
                jianglifenzhi = 'K'
                jine = 'J'
                zhandian = 'Q'
                kahao = 'A'
                shijian = 'B'
                pinzhong = 'E'
            }

            // 过滤掉积分卡加油和没有奖励积分的
            // if(val[yewuleixing]!== '加油'  || parseFloat(val[jianglifenzhi]) != 0 || parseFloat(val[jine]) == 0){
            // 过滤掉积分卡加油
            if(val[yewuleixing]!== '加油'  || parseFloat(val[jine]) == 0){
                return
            }
            // 站点_卡号_日期_柴油[汽油]
            let key = val[zhandian] + '_' + val[kahao] + '_' + val[shijian].substring(0,10)
            if(val[pinzhong].indexOf('汽油') !== -1){
                // 汽油
                key = key + '_汽油'
            }else{
                // 柴油
                key = key + '_柴油'
            }

            if( typeof result[key] === 'undefined' ){
                result[key] = 0
            }
            result[key] += parseFloat(val[jine])
        })
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

    let finalRes = {}
    files.forEach(function(file){
        mainWindow.webContents.send('read_path_one_by_one', file)
        var workbook = xlsx.readFile(file, {sheetRows:5100})
        var data = to_json(workbook)
        workbook = null
        if(Object.values(data).length === 0){
            return
        }
        for(var i in data){
            if( typeof finalRes[i] === 'undefined' ){
                finalRes[i] = data[i]
            }else {
                finalRes[i] = parseFloat(finalRes[i])+ parseFloat(data[i])
            }
        }
    })
    if(Object.values(finalRes).length === 0){
        mainWindow.webContents.send('read_path', '并没有统计到任何符合规则的信息')
        event.sender.send('read_path', '并没有统计到任何符合规则的信息')
        return
    }
    var _headers = ['站点','卡号', '日期','油品', '金额']
    var _data = []
    for(var i in finalRes){
        let ar = i.split('_')
        _data.push({
            '站点': ar[0],
            '卡号': ar[1],
            '日期': ar[2],
            '油品': ar[3],
            '金额': finalRes[i],
        })
    }
    var headers = _headers.map((v, i) => Object.assign({}, {v: v, position: String.fromCharCode(65+i) + 1 })).reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
    var data = _data.map((v, i) => _headers.map((k, j) => Object.assign({}, { v: v[k], position: String.fromCharCode(65+j) + (i+2) }))).reduce((prev, next) => prev.concat(next)).reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
    var output = Object.assign({}, headers, data);
    var outputPos = Object.keys(output);
    var ref = outputPos[0] + ':' + outputPos[outputPos.length - 1];
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
