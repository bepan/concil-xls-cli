const { app, BrowserWindow, ipcMain } = require('electron');
const url = require('url');
const path = require('path');
const workerpool = require('workerpool');

// Keep a global reference of the window object, if you don't, the window will
// be closed automatically when the JavaScript object is garbage collected.
let win;
var pool = workerpool.pool(path.join(__dirname, '../src/main.js'));

function createMainWindow () {
  // Create the browser window.
  win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      // devTools: false
    }
  })

  // and load the index.html of the app.
  win.loadURL(url.format({
    pathname: path.join(__dirname, 'index.html'),
    protocol: 'file',
    slashes: true
  }))

  // Open the DevTools.
  // win.webContents.openDevTools()

  // Remove default menu
  win.removeMenu();

  // Emitted when the window is closed.
  win.on('closed', () => {
    // Dereference the window object, usually you would store windows
    // in an array if your app supports multi windows, this is the time
    // when you should delete the corresponding element.
    win = null
  })
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.on('ready', createMainWindow)

// Quit when all windows are closed.
app.on('window-all-closed', () => {
  // On macOS it is common for applications and their menu bar
  // to stay active until the user quits explicitly with Cmd + Q
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', () => {
  // On macOS it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (win === null) {
    createWindow()
  }
})

// In this file you can include the rest of your app's specific main process
// code. You can also put them in separate files and require them here.
ipcMain.on('conciliate', function(e, args) {

  const moldedArgs = {
    file: args.baseFile,
    startFromCell: args.startFromCell,
    month: args.selectedMonth,
    year: args.selectedYear,
    outDir: args.outDir
  };

  // Run in a separate thread so it does not block the gui.
  pool.exec('conciliate', [moldedArgs])
    .then(function (result) {
      console.log('finish');
    })
    .catch(function (err) {
      console.error(err);
    })
    .then(function () {
      pool.terminate(); // terminate all workers when done
    });
});
