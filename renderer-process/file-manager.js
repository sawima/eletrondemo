const {shell} = require('electron')
const os = require('os')

const fileManagerBtn = document.getElementById('open-file-manager')

fileManagerBtn.addEventListener('click', (event) => {
//   shell.showItemInFolder(os.homedir())

    let filepath=event.currentTarget.dataset["filepath"]

    document.getElementById("loadspin").classList.remove('invisible')
    document.getElementById('showfilerow').classList.add('invisible')
    shell.showItemInFolder(filepath)
    setTimeout(() => {
      document.getElementById('showfilerow').classList.remove('invisible')
      document.getElementById("loadspin").classList.add('invisible')
    }, 6000)
})
