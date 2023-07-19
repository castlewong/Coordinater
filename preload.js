// preload.js
const { ipcRenderer } = require('electron');

window.addEventListener('DOMContentLoaded', () => {
  const chooseFileButton = document.getElementById('chooseFileButton');

  chooseFileButton.addEventListener('click', () => {
    ipcRenderer.send('chooseExcelFile');
  });

  ipcRenderer.on('excelData', (event, data) => {
    // Handle the data received from the main process
    if (data) {
      console.log('Excel Data:', data);
      // Display the data in your app or perform further processing
    } else {
      console.error('Error reading Excel file.');
      // Handle error if necessary
    }
  });
});
