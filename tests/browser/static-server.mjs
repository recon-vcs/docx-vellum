// Minimal static file server for the Playwright rendering tests.
// Serves the repository root so the harness can load dist/, node_modules/
// and tests/fixtures/ over HTTP. No third-party dependencies.
import { createServer } from 'node:http';
import { createReadStream } from 'node:fs';
import { stat } from 'node:fs/promises';
import { extname, join, normalize, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = fileURLToPath(new URL('../..', import.meta.url));
const PORT = 8765;

const MIME_TYPES = {
	'.html': 'text/html; charset=utf-8',
	'.js': 'text/javascript; charset=utf-8',
	'.mjs': 'text/javascript; charset=utf-8',
	'.cjs': 'text/javascript; charset=utf-8',
	'.css': 'text/css; charset=utf-8',
	'.json': 'application/json; charset=utf-8',
	'.map': 'application/json; charset=utf-8',
	'.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
	'.png': 'image/png',
	'.jpg': 'image/jpeg',
	'.svg': 'image/svg+xml',
	'.woff': 'font/woff',
	'.woff2': 'font/woff2',
};

const server = createServer(async (req, res) => {
	try {
		const url = new URL(req.url ?? '/', `http://127.0.0.1:${PORT}`);
		const relative = normalize(decodeURIComponent(url.pathname)).replace(/^([/\\])+/, '');
		const filePath = join(ROOT, relative);

		// Prevent path traversal outside the repository root.
		if (!filePath.startsWith(ROOT.endsWith(sep) ? ROOT : ROOT + sep)) {
			res.writeHead(403).end('Forbidden');
			return;
		}

		const info = await stat(filePath);
		if (!info.isFile()) {
			res.writeHead(404).end('Not found');
			return;
		}

		res.writeHead(200, {
			'Content-Type': MIME_TYPES[extname(filePath).toLowerCase()] ?? 'application/octet-stream',
			'Content-Length': info.size,
			'Cache-Control': 'no-store',
		});
		createReadStream(filePath).pipe(res);
	} catch {
		res.writeHead(404).end('Not found');
	}
});

server.listen(PORT, '127.0.0.1', () => {
	console.log(`static server listening on http://127.0.0.1:${PORT}/ (root: ${ROOT})`);
});
