const {shell} = require('electron')
const os = require('os')

const fileManagerBtn = document.getElementById('open-file-manager')

fileManagerBtn.addEventListener('click', (event) => {
//   shell.showItemInFolder(os.homedir())
console.log("filepath")

    let filepath=event.currentTarget.dataset["filepath"]
    console.log(filepath)
  shell.showItemInFolder(filepath)

//   event.currentTarget.classList.toggle('invisible')
})
