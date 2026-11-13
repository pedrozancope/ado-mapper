import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const routes = new Map([
  ['/', ['src/client/index.html', 'text/html; charset=utf-8']],
  ['/styles.css', ['src/client/styles.css', 'text/css; charset=utf-8']],
  ['/app.js', ['src/client/app.js', 'text/javascript; charset=utf-8']],
  ['/core.js', ['src/client/core.js', 'text/javascript; charset=utf-8']],
  ['/vendor/exceljs.js', ['node_modules/exceljs/dist/exceljs.min.js', 'text/javascript; charset=utf-8']],
  ['/vendor/jszip.js', ['node_modules/jszip/dist/jszip.min.js', 'text/javascript; charset=utf-8']]
]);

const server = http.createServer(async (request, response) => {
  const route = routes.get(request.url);
  if (request.method !== 'GET' || !route) {
    response.writeHead(404).end('Not found');
    return;
  }
  try {
    const data = await fs.readFile(path.join(root, route[0]));
    response.writeHead(200, {
      'Content-Type': route[1],
      'Cache-Control': 'no-store',
      'X-Content-Type-Options': 'nosniff',
      'Content-Security-Policy': "default-src 'self'; script-src 'self'; style-src 'self'; connect-src 'none'; img-src 'self' data:; object-src 'none'; base-uri 'none'; frame-ancestors 'none'",
      'Referrer-Policy': 'no-referrer'
    }).end(data);
  } catch (error) {
    console.error(`Unable to serve ${route[0]}: ${error.message}`);
    response.writeHead(500).end('Unable to load app files. Run npm install, then restart.');
  }
});

server.on('error', error => {
  console.error(`Cannot start local app: ${error.message}. Port 4173 must be available.`);
  process.exitCode = 1;
});
server.listen(4173, '127.0.0.1', () => {
  console.log('ADO Mapper: http://127.0.0.1:4173');
  console.log('Local files remain in your browser. No uploads or external connections.');
});
