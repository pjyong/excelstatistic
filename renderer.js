const ipc = require('electron').ipcRenderer

const selectDirBtn = document.getElementById('select-directory')
const resultDiv = document.getElementById('result')

selectDirBtn.addEventListener('click', function (event) {
    ipc.send('open_file_dialog')
})

ipc.on('accept_dir_path', function (event, choosePath) {
    if(choosePath !== null){
        appendHtml(choosePath, '您选择了文件:');
        appendHtml('请耐心等待...', '');
    }
})

ipc.on('read_path', function (event, data) {
    console.log('we get result')
    appendHtml(data, '读取结果:')
})

ipc.on('read_path_one_by_one', function (event, filename) {
    appendHtml(filename, '正在读取文件:');
})

function appendHtml(html, prefix){
    resultDiv.innerHTML = prefix + html + '<br/>' + resultDiv.innerHTML;
}

let myNotification = new Notification('Title', {
  body: 'Lorem Ipsum Dolor Sit Amet'
})

myNotification.onclick = () => {
  console.log('Notification clicked')
}
