# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project state

This is a newly scaffolded Python project (uv-based). The `src/excel_automation_script` package currently contains only a stub `main()` entry point, and `test` is an empty placeholder file (not a directory) with no test framework configured yet. There is no README content, lint config, or CI yet — treat architecture decisions as unmade rather than assuming existing conventions.

## Commands

This project uses [uv](https://docs.astral.sh/uv/) for dependency management and packaging, with `uv_build` as the build backend.

- Install/sync dependencies: `uv sync`
- Run the CLI entry point: `uv run excel-automation-script`
- Run any script/module through the venv: `uv run python -m excel_automation_script`
- Add a dependency: `uv add <package>`

## Architecture

- Package source lives under `src/excel_automation_script/` (src-layout).
- The console script `excel-automation-script` is defined in `pyproject.toml` under `[project.scripts]` and points to `excel_automation_script:main`.
- Requires Python >=3.12 (pinned via `.python-version` to 3.12).
- Runtime dependency: `pandas`.
