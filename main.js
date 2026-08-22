const { app, BrowserWindow } = require('electron');
const path = require('path');

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    title: "Hệ Thống Xếp Lớp - THCS Tân Xuân",
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  // Ẩn thanh menu trình duyệt cho giao diện giống phần mềm chuyên nghiệp
  win.setMenuBarVisibility(false);

  // Mở file index.html từ thư mục build khi chạy offline
  win.loadURL(`file://${path.join(__dirname, './build/index.html')}`);
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});