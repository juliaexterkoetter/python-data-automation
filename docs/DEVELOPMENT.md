# Development Workflow

## Assisted Development Process

The canonical process is defined in [Git and Pull Request Development Flow](#git-and-pull-request-development-flow). It requires planning before implementation, validation after changes, a separate review with an explicit `APPROVE` result before commit, and explicit user authorization before merge.

Implementation must not begin while behavior required by the task depends on an unresolved business decision.

## Engineering Guidelines

- Use type hints for public functions and wherever they improve correctness and understanding.
- Prefer small, focused functions with explicit responsibilities.
- Use clear English names for modules, functions, classes, and variables.
- Avoid premature abstractions and unnecessary architectural layers.
- Use logging instead of `print` for operational events.
- Raise or handle explicit exceptions at appropriate boundaries; do not suppress unexpected failures.
- Add docstrings when they explain contracts, assumptions, or behavior not already clear from the code.
- Do not add dependencies without a demonstrated need.
- Preserve traceability and make record classification explicit.
- Review generated code and the complete diff before considering work finished.
- Keep changes small enough to review and verify independently.

## Change Checklist

Before completing a code change:

1. Confirm that the behavior is supported by requirements or an accepted decision.
2. Assess effects on record accounting, traceability, errors, and output compatibility.
3. Add or update relevant tests.
4. Run the relevant test commands documented in `TESTING.md`.
5. Review the diff for unrelated changes, silent failure paths, and exposed secrets.
6. Update documentation if approved behavior or architecture changed.

## Git and Pull Request Development Flow

The standard project flow is:

```text
main
    -> create branch
    -> read context
    -> plan
    -> implement
    -> test
    -> review diff
    -> separate review
    -> REQUEST CHANGES?
        -> fix
        -> re-test
        -> re-review
    -> APPROVE
    -> commit
    -> push
    -> open pull request
    -> final review and checks
    -> user authorizes merge
    -> merge
    -> sync main
    -> delete branch
    -> next feature
```

### Main and Branch Scope

Never work directly on `main`. It represents reviewed and integrated code and documentation. Create each branch from an up-to-date `main` and limit it to one coherent objective.

Use these branch prefixes:

- `feature/<short-name>` for new functionality;
- `fix/<short-name>` for corrections;
- `docs/<short-name>` for documentation-only work;
- `test/<short-name>` for test-only work;
- `chore/<short-name>` for technical maintenance.

Names must be short, descriptive, and in English. Independent changes must not be mixed into the same branch. Divide large changes into smaller increments when practical.

Before implementation, confirm the branch, read the project context, establish the exact scope, identify ambiguities, and prepare a plan. Obtain approval before making an unresolved behavior or architectural decision.

### Implementation, Testing, and Review

Implement small, reviewable changes. Update documentation whenever approved behavior changes. Add or update tests with behavior changes, and give bugs found during review a regression test.

After implementation:

1. Run the relevant tests and checks.
2. Run `git diff --check`.
3. Review the complete diff against `main`.
4. Check for accidental scope expansion, secrets, credentials, local files, and generated artifacts.
5. Confirm that the implementation agrees with approved requirements and decisions.
6. Perform a code review as a separate step from implementation.

The implementing agent may assist with review, but the review result must remain an explicit, separate outcome. If it returns `REQUEST CHANGES`, correct the in-scope findings on the same branch, add or adjust tests, re-run validation, and perform another review. Repeat until the result is `APPROVE`.

Do not fix unrelated review findings without explaining why the scope must expand and obtaining approval when needed. Do not create a commit while the current review result is `REQUEST CHANGES`.

### Commit Messages

Commits must contain coherent, reviewable changes. Use short English messages with a conventional prefix:

- `feat:` for functionality;
- `fix:` for corrections;
- `docs:` for documentation;
- `test:` for tests;
- `chore:` for maintenance;
- `refactor:` for behavior-preserving restructuring.

Examples:

- `feat: add CSV input validation`
- `fix: reject CSV symlinks`
- `docs: document GitHub workflow`
- `test: cover filesystem access failures`

### Pull Request Titles

PR titles must prioritize human clarity and summarize the outcome without repeating the entire branch scope.

Good examples:

- `CSV import and validation`
- `Record normalization`
- `Order validation`
- `Duplicate handling`
- `Sales summary`
- `Excel report generation`

Avoid:

- `feat: add comprehensive CSV structural validation implementation`
- `feature/csv-input-validation`
- long titles that enumerate every technical detail in the branch.

### Pull Request Description

Every PR must use this structure:

```markdown
## Summary

Describe why this change exists.

## Changes

- Main change
- Supporting change

## Tests

- Tests executed
- Validation results

## Review notes

- Important decisions
- Known limitations
- Areas reviewers should inspect
```

The description must explain the change as real software engineering. It must not describe the project as only a portfolio example.

### Commit, Push, and PR Creation

After a separate review returns `APPROVE`, Codex may perform the following actions when requested:

1. Prepare and create the commit.
2. Push the branch to the remote.
3. Open the pull request.
4. Fill in its title and description using the project standards.

Do not open a pull request before the implementation is ready for review.

### Merge Protection and Strategy

Codex must never merge automatically. A PR may be merged only when:

1. A separate review has returned `APPROVE`.
2. All relevant checks pass.
3. There are no unreviewed changes.
4. The user explicitly authorizes the merge.

Technical mergeability is not authorization. Wait for explicit user approval.

Use a merge commit by default. Do not use squash merge or rebase merge without explicit authorization.

### After Merge

After an authorized merge:

1. Switch to `main`.
2. Run `git pull`.
3. Confirm that local `main` is current.
4. Delete the integrated local branch when safe.
5. Optionally remove the remote branch when that is part of the approved flow.
6. Create the next branch only from the updated `main`.
