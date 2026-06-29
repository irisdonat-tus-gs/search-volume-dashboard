const http = require('http');
const fs = require('fs');
const path = require('path');
const port = 3737;
const dir = path.join(__dirname, 'public');
const mime = {'.html':'text/html','.js':'application/javascript','.json':'application/json','.css':'text/css'};
http.createServer((req, res) => {
  const pathname = req.url.split('?')[0];
  const file = path.join(dir, pathname === '/' ? 'index.html' : pathname);
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); res.end('Not found'); return; }
    res.writeHead(200, {'Content-Type': mime[path.extname(file)] || 'text/plain'});
    res.end(data);
  });
}).listen(port, () => console.log('Server running on port ' + port));
