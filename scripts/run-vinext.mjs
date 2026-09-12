import { spawnSync } from 'node:child_process';
import path from 'node:path';
import process from 'node:process';

const root = path.resolve(import.meta.dirname, '..');
const localNode = path.join(root, 'node_modules', 'node', 'bin', process.platform === 'win32' ? 'node.exe' : 'node');
const cli = path.join(root, 'node_modules', 'vinext', 'dist', 'cli.js');
const result = spawnSync(localNode, [cli, ...process.argv.slice(2)], { cwd: root, stdio: 'inherit', env: process.env });
if (result.error) throw result.error;
process.exitCode = result.status ?? 1;
