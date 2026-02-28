# CLAUDE.md — AI Assistant Guide for kaizen

This file provides context and guidelines for AI assistants (Claude, Copilot, etc.) working in this repository.

---

## Project Overview

**Name:** kaizen
**Concept:** Kaizen (改善) is the Japanese philosophy of continuous improvement — small, incremental changes that compound over time.
**Status:** Early-stage / freshly initialized. No application code exists yet.

---

## Repository State

As of the last update to this file, the repository contains:

```
kaizen/
├── .git/           # Git metadata
├── README.md       # Project title only
└── CLAUDE.md       # This file
```

There are currently no source files, configuration files, dependencies, tests, or CI/CD pipelines. This is a blank slate.

---

## Git Workflow

### Branch Naming

- Feature branches: `feature/<short-description>`
- Bug fixes: `fix/<short-description>`
- Documentation: `docs/<short-description>`
- Claude-generated branches follow the pattern: `claude/<task-description>-<session-id>`

### Commit Style

Use concise, imperative commit messages:

```
Add user authentication module
Fix null pointer in payment processor
Update README with setup instructions
```

- Do **not** use past tense ("Added", "Fixed")
- Keep the subject line under 72 characters
- Add a blank line before the body if more detail is needed

### Pushing Changes

Always push with tracking:
```bash
git push -u origin <branch-name>
```

Branch names must start with `claude/` when working as an AI assistant on assigned tasks.

---

## Development Conventions (To Be Established)

As this project grows, conventions should be documented here. Recommended sections to add:

- **Language & Runtime** — e.g., Node.js 20+, Python 3.11+, Go 1.22+
- **Code Style** — linter/formatter in use (ESLint, Black, gofmt, etc.)
- **Testing** — framework and how to run tests
- **Build** — how to build the project locally
- **Environment** — required environment variables (use `.env.example`)
- **Dependencies** — how to install and update
- **Database** — migrations, seeding, local setup

---

## Instructions for AI Assistants

### General Principles

1. **Read before editing.** Always read a file in full before modifying it.
2. **Minimal changes.** Only change what is necessary. Avoid unnecessary refactors or unrelated cleanups.
3. **No speculative features.** Do not add error handling, abstractions, or features that are not requested.
4. **Security first.** Never introduce command injection, SQL injection, XSS, or other OWASP vulnerabilities.
5. **Avoid over-engineering.** Prefer simple, direct solutions over clever ones.

### File Creation Policy

- Do not create files unless explicitly required by the task.
- Prefer editing existing files over creating new ones.
- Do not create README or documentation files unless asked.

### Task Workflow

1. Check the current branch and git status before starting work.
2. Make changes on the designated development branch.
3. Commit with a clear, descriptive message.
4. Push to the remote branch using `git push -u origin <branch-name>`.

### What to Avoid

- Do not push to `master` or `main` without explicit user permission.
- Do not force-push (`--force`) without explicit user permission.
- Do not skip commit hooks (`--no-verify`).
- Do not amend published commits.
- Do not delete files, branches, or tables without user confirmation.
- Do not run destructive shell commands (`rm -rf`, `git reset --hard`, `git clean -f`) without confirmation.

---

## Adding to This File

As the project evolves, update this file to reflect:

- New frameworks, tools, or languages added
- Testing strategy and how to run the test suite
- CI/CD pipeline setup and how it works
- Environment setup instructions
- Architecture decisions and key design patterns
- Known gotchas or non-obvious conventions

This file is the primary source of truth for anyone (human or AI) coming into this codebase for the first time.
