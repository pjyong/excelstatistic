const ipc = require('electron').ipcRenderer

// 选择目录
const selectDirBtn = document.getElementById('select-directory')
selectDirBtn.addEventListener('click', function (event) {
    ipc.send('open_file_dialog')
})

// 操作结果显示Div
const resultDiv = document.getElementById('result')

function appendHtml(html, prefix){
    resultDiv.innerHTML = prefix + html + '<br/>' + resultDiv.innerHTML;
}

// 显示选择的目录信息
ipc.on('accept_dir_path', function (event, choosePath) {
    appendHtml(choosePath, '您选择了文件:');
    appendHtml('请耐心等待...', '');
})

ipc.on('read_path', function (event, data) {
    appendHtml(data, '读取结果:')
})

ipc.on('read_path_one_by_one', function (event, filename) {
    appendHtml(filename, '正在读取文件:');
})
