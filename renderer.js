// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const ipc = require('electron').ipcRenderer
// const BrowserWindow = require('electron').remote.BrowserWindow
// const url = require('url')
// const path = require('path')



const selectDirBtn = document.getElementById('select-directory')
const resultDiv = document.getElementById('result')

selectDirBtn.addEventListener('click', function (event) {
    ipc.send('open_file_dialog')
})

ipc.on('accept_dir_path', function (event, choosePath) {
    if(choosePath !== null){
        appendHtml(choosePath, '您选择了目录:');
        appendHtml('请耐心等待...', '');
        /*
        setTimeout(function(){
            ipc.send('choose_dir', path)
        }, 1000)
        */

    }
})

ipc.on('read_path_to_main', function (event, data) {
    console.log('we get result')
    appendHtml(data, '读取结果:')
})

ipc.on('read_path_one_by_one', function (event, filename) {
    appendHtml(filename, '正在读取文件:');
})

function appendHtml(html, prefix){
    resultDiv.innerHTML = prefix + html + '<br/>' + resultDiv.innerHTML;
}
