# KhmerTools

KhmerTools is a responsive Node.js and Express web application for PDF conversion, OCR, image processing, Gemini chat, QR creation, and safe direct public-media downloads. The frontend is vanilla HTML, CSS, and JavaScript.

## Requirements

- Node.js 20 or newer
- npm
- Poppler for PDF page rendering and optional PDF compression
- Internet access on first OCR use so Tesseract.js can obtain Khmer/English language data

## Install and run

```bash
npm install
npm start
```

Open `http://localhost:3000`.

Copy `.env.example` to `.env` and change values for your environment. Never commit `.env`.

## Gemini setup

1. Create an API key in Google AI Studio.
2. Set `GEMINI_API_KEY` in `.env`.
3. Optionally set `GEMINI_MODEL` (default: `gemini-3.6-flash`).

Gemini requests are sent by the Express server. The browser never receives the API key. If the key is absent, the AI route returns a clear configuration error while all non-AI tools continue working.

## Poppler setup

PDF-to-images, scanned-PDF OCR, and scanned-PDF-to-Word need `pdftoppm`. PDF compression additionally needs `pdftocairo`.

- Windows: install a Poppler build, then set absolute paths for `PDFTOPPM_PATH` and `PDFTOPDF_PATH` in `.env`.
- macOS: `brew install poppler`
- Debian/Ubuntu: `sudo apt-get install poppler-utils`

When the commands are on `PATH`, the defaults work without changes.

## Optional yt-dlp setup

`YTDLP_PATH` is reserved for a future official/public media integration. The current downloader intentionally supports only direct, publicly accessible image/audio/video URLs. It does not scrape platform pages or bypass DRM, logins, private content, paywalls, or other protections. If yt-dlp is later integrated, restrict it to lawful public content and keep URL/redirect validation in place.

## API routes

- `GET /api/status`
- `POST /api/ocr`
- `POST /api/pdf-ocr`
- `POST /api/pdf-to-word`
- `POST /api/pdf-ocr-to-word`
- `POST /api/ocr-to-excel`
- `POST /api/pdf-to-images`
- `POST /api/pdf-merge`, `/api/pdf-split`, `/api/pdf-rotate`, `/api/pdf-compress`
- `POST /api/image-convert`, `/api/image-compress`, `/api/image-resize`, `/api/image-to-pdf`
- `POST /api/qr`
- `POST /api/ai/chat`
- `POST /api/media/download`

Uploads are size-limited, signature-checked, sanitized, and processed from memory or temporary directories. Temporary processing directories are deleted after each job. Media URLs are restricted to public HTTP/HTTPS hosts and each redirect is revalidated against local/private address ranges.

## Deployment (Render or similar)

- Build command: `npm install`
- Start command: `npm start`
- Runtime: Node.js 20+
- Add environment variables from `.env.example`
- Install Poppler in the host image if PDF rendering/compression is required
- Set `ALLOWED_ORIGIN` to the production origin

The server uses the platform `PORT` and listens on `0.0.0.0`.

## Production notes

- Replace `khmertools.example` in `public/index.html`, `public/robots.txt`, and `public/sitemap.xml` with the real domain.
- Configure Stripe or another provider before enabling paid plan buttons. No payment endpoint is included.
- Add authentication and persistent quotas before offering accounts or commercial API access.
- Use a job queue/worker pool for high-volume OCR deployments.
