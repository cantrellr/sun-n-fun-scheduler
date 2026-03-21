# Contributing

## Scope

This repository is a Windows PowerShell scheduling prototype with operational
data concerns. Contributions should prioritize correctness, clarity, and low-risk
change sets.

## Before You Change Anything

1. Read `README.md`.
2. Read `DEVELOPMENT.md`.
3. Understand whether your change affects the GUI, exports, or data files.
4. Avoid changing seed data unless the baseline schedule itself is intentionally
   being updated.

## Contribution Guidelines

- Keep pull requests focused and easy to review.
- Do not mix documentation, data changes, and behavior changes without a reason.
- Preserve Windows compatibility.
- Preserve the non-Excel export path.
- Prefer small refactors over large rewrites in this prototype stage.

## Data Handling Rules

- Treat volunteer contact information as sensitive.
- Do not paste raw schedule data into public issues or screenshots.
- Avoid committing generated exports unless the repository explicitly needs a
  sample artifact updated.

## Manual Validation Expectations

Validate the parts you touch.

### If you change launch or UI behavior

- Start the app using `Start-SunNFunScheduler.cmd`.
- Confirm the day list loads.
- Confirm a day can be opened.

### If you change export behavior

- Run `Export-SunNFunSchedule.ps1`.
- Verify CSV/HTML outputs are generated.
- Verify Excel output if Excel is available and your change touched that path.

### If you change data reset or data loading logic

- Run `Reset-SeedData.ps1`.
- Confirm the working JSON file is restored correctly.

## Pull Request Notes

Include the following in your pull request description:

- what changed
- why it changed
- any risks or operational implications
- what you manually validated
- whether any seed or working data files changed

## Style

- Match the existing repository style.
- Keep comments brief and only where they clarify non-obvious logic.
- Prefer explicit, readable PowerShell over dense one-liners.

## Documentation

Update documentation when behavior, workflow, or assumptions change.

At minimum, update `README.md` or `DEVELOPMENT.md` when a contribution changes:

- required prerequisites
- launch flow
- export behavior
- data format expectations
- known limitations
