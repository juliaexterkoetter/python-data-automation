# Development Workflow

## Assisted Development Process

```text
requirements
    -> analysis
    -> clarify ambiguities
    -> approve decisions
    -> update documentation
    -> plan implementation
    -> implement a small change
    -> test
    -> review the diff
    -> update documentation if necessary
    -> commit
```

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
