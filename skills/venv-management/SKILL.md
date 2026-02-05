# venv-management / SKILL.md

## Purpose

This project uses **`uv`** for **all Python virtual environment and dependency management**.

The goal is to:
- Keep dependencies isolated
- Avoid dependency conflicts
- Ensure fully reproducible setups
- Prevent accidental use of global Python packages

This document defines the **only allowed way** to manage Python environments and packages in this project.

---

## Tooling Choice (MANDATORY)

- ✅ **uv** — the only allowed package & environment manager
- ❌ pip
- ❌ pipenv
- ❌ poetry
- ❌ conda
- ❌ global Python installs

If you suggest or use any package-management command, it **must** be a `uv` command.

---

## Project Structure & Environment Model

- The virtual environment is managed automatically by **uv**
- Installed dependencies live inside the project’s **`.uv/` directory**
- Project metadata and dependencies are defined in **`pyproject.toml`**

There is **no reliance on system-wide Python packages**.

---

## Commands You MUST Use

### Add a new dependency
Install a library and register it in `pyproject.toml`:

```bash
uv add <library>
```

---

### Install / sync all dependencies
After cloning the repository or pulling changes:

```bash
uv sync
```

---

### Run commands inside the environment
Any Python execution must happen through `uv`:

```bash
uv run <command>
```

Examples:
```bash
uv run python main.py
uv run pytest
uv run python -m my_module
```

---

## Dependency Rules

- Prefer the **Python standard library** whenever possible
- Do **not** add dependencies unless clearly justified
- Never install packages globally
- Never manually edit `.uv/` contents
- Never bypass `uv` to “quickly test something”

---

## Forbidden Behavior

- Running `pip install`
- Activating a venv manually
- Mixing package managers
- Installing dependencies without updating `pyproject.toml`

---

## Mental Model

> “If it’s not in `pyproject.toml`, it doesn’t exist.”

> “If it wasn’t installed with `uv`, it’s not allowed.”
