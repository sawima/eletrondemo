const {ipcMain, dialog} = require('electron')

const {spawn}=require('child_process')

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
      event.sender.send('selected-directory', files)
      console.log(files)
      //Todo: spawn go process to handle data transformation
      // spawn('./bin/main')
      tbexcel.readWorkbook(files[0])
    }
  })
})
