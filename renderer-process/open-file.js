const {ipcRenderer} = require('electron')

const selectDirBtn = document.getElementById('select-directory')

selectDirBtn.addEventListener('click', (event) => {
  console.log("open file")
  ipcRenderer.send('open-file-dialog')
})

ipcRenderer.on('selected-directory', (event, result) => {
    if(result){
      if(result.success){
        let fileManagerBtn=document.getElementById('open-file-manager')
        fileManagerBtn.setAttribute("data-filepath",result.target)
        document.getElementById('selectfilerow').classList.add('invisible')
        console.log("xxxxxx")
        document.getElementById("loadspin").classList.remove('invisible')
        setTimeout(function(){
          document.getElementById("loadspin").classList.add('invisible')
          document.getElementById('showfilerow').classList.remove('invisible')
        },6000)
        document.getElementById('selected-file').innerHTML = ""
      } else{
        document.getElementById('selected-file').innerHTML = "faild, please retry"
      }
    }
})

document.getElementById("returnlink").addEventListener('click',(e)=>{
  e.preventDefault()
  document.getElementById('selectfilerow').classList.remove('invisible')
  document.getElementById('showfilerow').classList.add('invisible')
})
