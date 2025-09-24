# VVS Local App

The VVS Local App is a self-hosted, single-binary application that mirrors the existing Google Apps Script flows in a local environment. It serves static HTML/JS assets, exposes a JSON API, and persists data to SQLite by default so the system can run entirely offline. The same interfaces can later be pointed at Postgres and object storage when we deploy to a VPS.

## Project Goals

* Ship a deterministic, testable monolith without SaaS lock-in.
* Serve static UI assets from the same binary that powers the JSON API and job scheduler.
* Persist business data to SQLite locally, with a planned upgrade path to Postgres + MinIO.
* Mirror the semantics and naming conventions established in the existing Apps Script implementation.

## Repository Layout

```
.
├── app/                # compiled binary output (ignored in git)
├── backup/             # zipped database + file backups
├── cmd/vvsapp/         # application entry point
├── config/             # application configuration (YAML + env overrides)
├── files/              # uploaded binary assets (YYYY/MM/...)
├── internal/           # packages for config, logging, db, auth, server
├── local.db            # SQLite database (ignored in git)
├── logs/               # JSON logs written by the app
├── web/                # static UI (HTML/CSS/JS)
└── ...
```

## Getting Started

1. **Install Go 1.21+** and ensure `$GOPATH/bin` is on your PATH.
2. **Install build dependencies:**
   * The project uses Go modules and the pure-Go `modernc.org/sqlite` driver, so no CGO toolchain is required.
3. **Configure environment:**
   * Copy `.env.example` to `.env` (or export variables directly) and set the admin credentials and JWT secret.
   * Review `config/app.yaml` for defaults such as server address and database path.
4. **Build the binary:**
   ```bash
   go build -o app/vvsapp ./cmd/vvsapp
   ```
5. **Run migrations & start the server:**
   ```bash
   ./app/vvsapp
   ```
   The application will load configuration, run migrations, seed the admin user, and start an HTTP server.
6. **Check health:**
   * `GET http://localhost:8080/api/health` should return `{"status":"ok","db":"ok",...}`.
   * `POST http://localhost:8080/api/auth/login` with the seeded admin credentials returns a JWT.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for coding style, commit hygiene, and review expectations.

## Licensing

This repository is proprietary and intended for internal use.
