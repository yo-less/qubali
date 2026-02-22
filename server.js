const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const ROOT = __dirname;
const DATA_FILE = path.join(ROOT, 'data', 'entries.json');

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

function send(res, code, body, type = 'text/plain; charset=utf-8') {
  res.writeHead(code, { 'Content-Type': type });
  res.end(body);
}

function serveFile(reqPath, res) {
  const safePath = path.normalize(reqPath).replace(/^\.+/, '');
  const filePath = path.join(ROOT, safePath === '/' ? '/index.html' : safePath);
  if (!filePath.startsWith(ROOT)) return send(res, 403, 'Forbidden');
  fs.readFile(filePath, (err, data) => {
    if (err) return send(res, 404, 'Not Found');
    const ext = path.extname(filePath).toLowerCase();
    send(res, 200, data, MIME[ext] || 'application/octet-stream');
  });
}

const server = http.createServer((req, res) => {
  if (req.method === 'GET' && req.url === '/api/state') {
    fs.readFile(DATA_FILE, (err, data) => {
      if (err) return send(res, 500, JSON.stringify({ error: 'read_failed' }), MIME['.json']);
      send(res, 200, data, MIME['.json']);
    });
    return;
  }

  if (req.method === 'POST' && req.url === '/api/state') {
    let body = '';
    req.on('data', (chunk) => { body += chunk; if (body.length > 5_000_000) req.destroy(); });
    req.on('end', () => {
      try {
        const parsed = JSON.parse(body || '{}');
        fs.writeFile(DATA_FILE, JSON.stringify(parsed, null, 2), (err) => {
          if (err) return send(res, 500, JSON.stringify({ error: 'write_failed' }), MIME['.json']);
          send(res, 200, JSON.stringify({ ok: true, file: '/data/entries.json' }), MIME['.json']);
        });
      } catch {
        send(res, 400, JSON.stringify({ error: 'invalid_json' }), MIME['.json']);
      }
    });
    return;
  }

  serveFile(req.url.split('?')[0], res);
});

server.listen(PORT, () => {
  console.log(`Qubali app running on http://localhost:${PORT}`);
  console.log(`Data file: ${DATA_FILE}`);
});
