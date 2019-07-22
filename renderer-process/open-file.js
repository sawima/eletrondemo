const {ipcRenderer} = require('electron')

const selectDirBtn = document.getElementById('select-directory')

selectDirBtn.addEventListener('click', (event) => {
  console.log("open file")
  ipcRenderer.send('open-file-dialog')
})

ipcRenderer.on('selected-directory', (event, result) => {
  // document.getElementById('selected-file').innerHTML = `${path}`
  // document.getElementById('selected-file').innerHTML = `${path}`
  console.log(result)
    if(result){
      if(result.success){
        let fileManagerBtn=document.getElementById('open-file-manager')
        fileManagerBtn.setAttribute("data-filepath",result.target)
        fileManagerBtn.classList.toggle("invisible")
        document.getElementById('selected-file').innerHTML = ""
      } else{
        document.getElementById('selected-file').innerHTML = "faild, please retry"
      }
    }

})
