const {ipcMain, dialog} = require('electron')

ipcMain.on('open-file-dialog', (event) => {
  dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [
      { name: 'excel', extensions: ['xlsx'] }
    ]
  }, (files) => {
    if (files) {
      event.sender.send('selected-directory', files)

      //Todo: spawn go process to handle data transformation
      
    }
  })
})
