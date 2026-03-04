# Large Upload Setup (250MB-500MB ZIP)

This project is now set up for split deployment:

- Frontend: Vercel (static site)
- Backend API: separate Node service (Express) that handles large uploads

## 1) Deploy API service

Use any Node host where you control request size limits (VM, container platform, etc.).

Required environment variables for API:

- `API_ONLY=true`
- `PORT=5000` (or platform-provided port)
- `CORS_ORIGIN=https://<your-vercel-domain>`

Build and run:

```bash
npm ci
npm run build:api
npm run start:api
```

Health check:

- `GET /api/health`

Extract endpoint:

- `POST /api/extract` (multipart form-data with `files`)

## 2) Point Vercel frontend to API

In Vercel Project Settings -> Environment Variables:

- `VITE_EXTRACT_API_URL=https://<your-api-domain>/api/extract`

Redeploy the Vercel project after setting the variable.

## 3) Request-size limits

This app is configured for up to ~550MB request payloads in app logic, but your hosting/proxy must also allow it.

If you run behind Nginx, set:

```nginx
client_max_body_size 550M;
```

If your host has a stricter hard limit, choose a host/plan that supports 500MB uploads.
