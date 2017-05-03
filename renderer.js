// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.
const ipc = require('electron').ipcRenderer

const selectDirBtn = document.getElementById('select-directory')

selectDirBtn.addEventListener('click', function (event) {
    ipc.send('open-file-dialog')
})

ipc.on('result', function (event, data) {
    var resDiv = document.getElementById('selected-file')
    if(data == ''){
        document.getElementById('selected-file').innerHTML = `请选择有文件的目录`
    }else{
        document.getElementById('selected-file').innerHTML = data


    }
})
