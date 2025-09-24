# Contributing Guide

We ship a deterministic, auditable system. All contributions must keep those guarantees intact.

## Workflow

1. **Discuss before large changes.** Open an issue or start a thread describing the desired behavior and rationale.
2. **Small, focused commits.** Each commit should compile, pass tests, and describe _why_ the change exists.
3. **Code review required.** A second reviewer must sign off before merging to `main`.
4. **No direct pushes to `main`.** Use pull requests for every change.

## Coding Standards

* **Go version:** 1.21 or higher.
* **Formatting:** `gofmt` for Go, EditorConfig settings for everything else.
* **Linting:** Run `go test ./...` and any repo-specific checks before submitting a PR.
* **Error handling:** Return contextual errors; avoid panics outside of `main()`.
* **Logging:** Structured JSON logs via the shared logging package.
* **Security:** Never store plaintext secrets. Use env vars or secret managers.

## Commit Messages

* Imperative tense (e.g., "Add health endpoint").
* Reference issues or ADRs when relevant.
* Include context for operational changes (migrations, config).

## Pull Request Checklist

* [ ] Tests added or updated.
* [ ] Migrations are idempotent and reversible (or documented otherwise).
* [ ] Configuration changes documented in README or inline comments.
* [ ] No secrets, API keys, or customer data checked in.
* [ ] Health checks and diagnostics still succeed locally.

Thanks for helping keep the project healthy!
