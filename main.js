const {
    app,
    Menu,
    BrowserWindow
} = require("electron");

require("electron-reload")(__dirname, {
    // Note that the path to electron may vary according to the main file
    electron: require(`${__dirname}/node_modules/electron`)

});
// const custom_menu = require('./menu');

// const { menu } = require ( );

let win;

function createWindow() {
    // Create the browser window.
    win = new BrowserWindow({
        width: 800,
        height: 600
    });

    // and load the index.html of the app.
    win.loadURL(`file://${__dirname}/app/src/onboard.html`);


    // Menu.setApplicationMenu(custom_menu);

    // Open the DevTools.
    // win.webContents.openDevTools();

    // Emitted when the window is closed.
    win.on("closed", () => {

        win = null;
    });
}


app.on("ready", createWindow);

// Quit when all windows are closed.
app.on("window-all-closed", () => {

    if (process.platform !== "darwin") {
        app.quit();
    }
});

app.on("activate", () => {

    if (win === null) {
        createWindow();
    }
});