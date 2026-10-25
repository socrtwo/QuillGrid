// electron-builder configuration for the desktop releases.
// Windows: NSIS installer (x64, arm64) and a portable .exe (x64).
// macOS: universal (Apple silicon + Intel) .dmg and .zip, ad-hoc signed.
// Linux: AppImage and .deb for x64, arm64 (Raspberry Pi OS 64-bit, ChromeOS Linux on ARM) and
// armv7l (Raspberry Pi OS 32-bit).
/** @type {import('electron-builder').Configuration} */
module.exports = {
  appId: 'io.github.socrtwo.quillgrid',
  productName: 'Anne\'s Grid',
  executableName: 'QuillGrid', // the binary keeps its old name; the product (display) name is Anne's Grid
  copyright: 'Anne\'s Grid contributors, MIT License',
  directories: { app: 'build/desktop-app', output: 'release', buildResources: 'resources' },
  artifactName: 'QuillGrid-${version}-${os}-${arch}.${ext}',
  asar: true,
  // Everything the app needs is bundled into dist/ by Vite; ship no node_modules.
  files: ['dist/**/*', 'electron/**/*', 'package.json', '!**/node_modules/**'],
  icon: 'resources/icon-1024.png',
  fileAssociations: [
    { ext: 'xlsx', name: 'Excel Workbook', role: 'Editor', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
    { ext: 'xlsm', name: 'Excel Macro-Enabled Workbook', role: 'Editor', mimeType: 'application/vnd.ms-excel.sheet.macroEnabled.12' },
    { ext: 'xltx', name: 'Excel Template', role: 'Editor', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template' },
    { ext: 'xls', name: 'Excel 97-2003 Workbook', role: 'Viewer', mimeType: 'application/vnd.ms-excel' },
    { ext: 'ods', name: 'OpenDocument Spreadsheet', role: 'Editor', mimeType: 'application/vnd.oasis.opendocument.spreadsheet' },
    { ext: 'fods', name: 'Flat OpenDocument Spreadsheet', role: 'Editor', mimeType: 'application/vnd.oasis.opendocument.flat-spreadsheet' },
    { ext: 'csv', name: 'Comma-Separated Values', role: 'Editor', mimeType: 'text/csv' },
    { ext: 'tsv', name: 'Tab-Separated Values', role: 'Editor', mimeType: 'text/tab-separated-values' },
    { ext: 'qgw', name: 'Anne\'s Grid Workbook', role: 'Editor', mimeType: 'application/x-quillgrid' },
    { ext: 'qgwe', name: 'Encrypted Anne\'s Grid Workbook', role: 'Editor', mimeType: 'application/x-quillgrid-encrypted' },
    { ext: 'qlw', name: 'QuillGrid Workbook (QuillLattice)', role: 'Editor', mimeType: 'application/x-quillgrid' }, // rename:keep — QuillGrid was named QuillLattice; its .qlw/.qlwe files still open
    { ext: 'qlwe', name: 'Encrypted QuillGrid Workbook (QuillLattice)', role: 'Editor', mimeType: 'application/x-quillgrid-encrypted' }, // rename:keep — QuillGrid was named QuillLattice; its .qlw/.qlwe files still open
  ],
  win: {
    target: [{ target: 'nsis', arch: ['x64', 'arm64'] }, { target: 'portable', arch: ['x64'] }],
  },
  nsis: { oneClick: false, perMachine: false, allowToChangeInstallationDirectory: true, artifactName: 'QuillGrid-${version}-windows-${arch}-setup.${ext}' },
  portable: { artifactName: 'QuillGrid-${version}-windows-portable.${ext}' },
  mac: {
    target: [{ target: 'dmg', arch: ['universal'] }, { target: 'zip', arch: ['universal'] }],
    category: 'public.app-category.productivity',
    extendInfo: { CFBundleDisplayName: 'Anne\'s Grid' },
    // Ad-hoc signature: runs on Apple silicon without a paid Developer ID, but is not notarized,
    // so Gatekeeper asks once (docs/INSTALL.md). Hardened runtime would reject Electron's own signature.
    identity: '-',
    hardenedRuntime: false,
    gatekeeperAssess: false,
  },
  linux: {
    target: [{ target: 'AppImage', arch: ['x64', 'arm64', 'armv7l'] }, { target: 'deb', arch: ['x64', 'arm64', 'armv7l'] }],
    category: 'Office;Spreadsheet',
    synopsis: 'Spreadsheet with a complete, customizable command library',
    description: 'Anne\'s Grid is an open-source spreadsheet that opens and saves Excel (.xlsx, .xls), OpenDocument (.ods) and CSV files. Every feature is a command you can put on the ribbon, the Quick Access Bar, a shortcut or a macro.',
    maintainer: 'socrtwo <socrtwo@users.noreply.github.com>',
    executableName: 'quillgrid',
    syncDesktopName: true,
  },
  publish: null,
};
