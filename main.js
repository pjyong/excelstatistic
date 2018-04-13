const electron = require('electron')
const app = electron.app
const BrowserWindow = electron.BrowserWindow
const path = require('path')
const url = require('url')
let mainWindow

function createWindow () {
    mainWindow = new BrowserWindow({width: 800, height: 600})
    mainWindow.loadURL(url.format({
        pathname: path.join(__dirname, 'index.html'),
        protocol: 'file:',
        slashes: true
    }))

    // 是否开启调试工具
    // mainWindow.webContents.openDevTools()
    mainWindow.on('closed', function () {
        mainWindow = null
    })
}
app.on('ready', createWindow)
app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') {
        app.quit()
    }
})

app.on('activate', function () {
    if (mainWindow === null) {
        createWindow()
    }
})

// 开始业务代码
const ipc = require('electron').ipcMain
const dialog = require('electron').dialog
const xlsx = require('xlsx')
const fs = require('fs')

// 遍历目录，获取所有符合规则的文件
function getAllFiles(root){
    var res = [] , files = fs.readdirSync(root);
    files.forEach(function(file){
        var pathname = root+'/'+file, stat = fs.lstatSync(pathname);
        if(!stat.isDirectory()){
            // 过滤积分和财务统计的excel
            if( /^.*\.xls$/.test(pathname) &&  pathname.indexOf('财务统计')=== -1 && pathname.indexOf('jf')=== -1 ){
                res.push(pathname);
            }
        }else{
            // 遍历目录
            if(pathname.indexOf('jf')=== -1 && pathname.indexOf('统计')=== -1){
                res = res.concat(getAllFiles(pathname));
            }
        }
    });
    return res
}

// 从Excel提取出相应的数据
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
            if(val[yewuleixing]!== '加油'  || parseFloat(val[jine]) == 0){
                return
            }
            // 站点_卡号_日期_[柴油OR汽油]
            let key = val[zhandian] + '_' + val[kahao] + '_' + val[shijian].substring(0,10)
            if(val[pinzhong].indexOf('汽油') !== -1){
                key = key + '_汽油'
            }else{
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
        if(typeof choosePath !== 'undefined' ){
            // 发个消息给前台显示目录信息
            event.sender.send('accept_dir_path', choosePath)
            // 新建个窗口执行后台任务，用户不可见
            var win = new BrowserWindow({
                width: 400, height: 225, show: false
            })
            win.loadURL(url.format({
                pathname: path.join(__dirname, 'invisible.html'),
                protocol: 'file:',
                slashes: true
            }))
            win.webContents.on('did-finish-load', function () {
                win.webContents.send('bg_accept_path', choosePath)
            })
        }
    })
})

// 处理后台进程发的请求
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
