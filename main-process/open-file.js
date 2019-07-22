const {ipcMain, dialog} = require('electron')

const {spawn}=require('child_process')
const path=require("path")

// const {tbexcel}=require('../tbexcel/tbexcel')
const tbexcel=require('../tbexcel/tbexcel')

ipcMain.on('open-file-dialog', (event) => {
  dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [
      { name: 'excel', extensions: ['xlsx'] }
    ]
  }, (files) => {
    if (files) {
      // event.sender.send('selected-directory', files)
      // event.sender.send('selected-directory', path.dirname(files[0]))
      event.sender.send('selected-directory', files[0])
      //Todo: spawn go process to handle data transformation
      // spawn('./bin/main')
      let targetFile=tbexcel.readWorkbook(files[0])
      console.log(targetFile)
      event.sender.send('selected-directory', targetFile)

      // if(transformResult.success==true){
      //   event.sender.send('selected-directory', transformResult.target)
      // } else{
      //   event.sender.send('selected-directory', "failed")
      // }
    }
  })
})
