# Agent Instructions

Before working on this project:

1. Read `docs/PROJECT.md`.
2. Read `docs/requirements.md`.
3. Read `docs/BUSINESS_RULES.md`.
4. Read `docs/ARCHITECTURE.md`.
5. Read `docs/DECISIONS.md`.
6. Consult `docs/DEVELOPMENT.md` before changing code.
7. Consult `docs/TESTING.md` when implementing or fixing behavior.
8. Consult `docs/TASKS.md` for the current work status.

## Permanent Rules

- Never invent missing business rules. Document ambiguities and request approval.
- Never discard data silently. Preserve data traceability throughout processing.
- Do not change documented architectural decisions without explaining the impact and requesting approval.
- Analyze the impact before relevant changes.
- Keep changes small and reviewable.
- Add or update tests whenever behavior changes, and run the relevant tests afterward.
- Never commit secrets or credentials.
- Keep code, documentation, comments, tests, logs, and application messages in English.
- Communicate with the project owner in Portuguese.
- Update the documentation when an approved decision changes system behavior.

## Git and GitHub Workflow

### Branches

- Never work directly on `main`. It must contain only reviewed and integrated work.
- Use a separate branch for each independent change.
- Name branches in English with a short descriptive name: `feature/<short-name>`, `fix/<short-name>`, `docs/<short-name>`, `test/<short-name>`, or `chore/<short-name>`.
- Examples: `feature/csv-input-validation`, `feature/data-normalization`, `fix/source-row-calculation`, and `docs/project-foundation`.

### Before Implementation

1. Confirm the active branch.
2. Read `AGENTS.md` and the relevant documentation.
3. Confirm the branch scope and identify ambiguities.
4. Present a plan before relevant changes.
5. Wait for approval when behavior or architecture requires a decision.

### Before Commit

1. Run relevant tests and project checks.
2. Run `git diff --check`.
3. Review the complete diff against `main` for accidental scope expansion.
4. Check for secrets, credentials, local files, and accidental artifacts.
5. Confirm that implementation and approved documentation agree.
6. Do not commit while a review result is `REQUEST CHANGES`.

### Commits and Pull Requests

- Use coherent, reviewable commits with short English messages prefixed by `feat:`, `fix:`, `docs:`, `test:`, `chore:`, or `refactor:`.
- Examples: `feat: add CSV input validation`, `fix: reject CSV symlinks`, `docs: document GitHub workflow`, and `test: cover filesystem access failures`.
- Use short, human-readable PR titles such as `CSV import and validation`, `Data normalization`, `Duplicate detection`, or `Excel report generation`; do not use branch names, commit-message syntax, or exhaustive technical titles.
- Use the PR description structure defined in `docs/DEVELOPMENT.md`: `Summary`, `Changes`, `Tests`, and `Review notes`.
- Describe the work as real software engineering; never characterize the project as only a portfolio example.

### Review and Merge

1. After implementation, test, review the diff, and perform a separate code review.
2. On `REQUEST CHANGES`, fix in-scope findings on the same branch, add or update tests, and re-review until `APPROVE`.
3. After `APPROVE`, commit, push, and open the PR only when requested.
4. Never merge automatically. Merge only after separate `APPROVE`, passing checks, no unreviewed changes, and explicit user authorization.
5. Use a merge commit by default; do not use squash or rebase merge without explicit authorization.
6. After an authorized merge, switch to `main`, pull, confirm it is current, safely remove the integrated local branch, optionally remove the remote branch when approved, and create the next branch from updated `main`.

See `docs/DEVELOPMENT.md` for the detailed development, review, PR, and post-merge process.
