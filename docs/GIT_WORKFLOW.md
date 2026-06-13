# Git Workflow

Repository:

```text
https://github.com/mxrzhang-afk/BERTToolkitXLTemplate.git
```

## Branches

- `main`: known-good toolkit and template state only.
- `feature/<tab-or-action>`: one tab/action definition at a time, for example `feature/input-gnpi-charts`.
- `fix/<issue>`: repair-only changes, for example `fix/template-corruption`.

## Workbook Rules

- Keep `.xlsm` templates in Git as binary files.
- Do not edit workbook XML directly unless there is no safer option.
- Prefer changing VBA/R code and inserting generated outputs at runtime.
- Before touching a workbook template, commit the current known-good file.
- If Excel reports file repair/corruption, discard the branch or restore the `.xlsm` from Git instead of creating ad hoc backup files.

## Tab Action Workflow

1. Create a branch from `main`.
2. Define the tab action contract in `r/tools/xl_pricing_tool/schema.yml`.
3. Implement R/VBA changes.
4. Run local smoke tests for the action.
5. Validate the `.xlsm` opens in Excel before merge.
6. Commit only after validation is approved.
7. Open a PR and merge to `main`.

## Useful Commands

```bash
git checkout main
git pull
git checkout -b feature/input-gnpi-charts
git status
git add .
git commit -m "Implement Input_GNPI chart action"
git push -u origin feature/input-gnpi-charts
```

To recover a corrupted workbook from the last commit:

```bash
git restore excel/templates/Template_XOL_Pricing_Tool.xlsm
```

