const {app, BrowserWindow} = require("electron");
let ejs = require('ejs-electron');
ejs.data({
    rows: 50,
    cols: 50,
});
let win;
function createwindow(){
    win = new BrowserWindow({
        width: 800,
        height: 600,
        show: false,
        webPreferences: {
            nodeIntegration: true,
            enableRemoteModule: true
        }
    });
}

app.whenReady().then(function(){
    createwindow();
    win.loadFile('index.ejs').then(function(){
        win.maximize();
        win.show();
    });
});