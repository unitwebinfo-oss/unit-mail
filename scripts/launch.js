// Launcher que remove ELECTRON_RUN_AS_NODE do ambiente antes de
// invocar o electron-vite. Sem isso, o Electron roda como Node puro
// e require('electron') não expõe a API do app.
delete process.env.ELECTRON_RUN_AS_NODE;

const { spawnSync } = require('child_process');
const { join, dirname } = require('path');
const bin = join(dirname(require.resolve('electron-vite/package.json')), 'bin', 'electron-vite.js');
const result = spawnSync(process.execPath, [bin, ...process.argv.slice(2)], {
  stdio: 'inherit',
  env: process.env,
});
process.exit(result.status ?? 0);
