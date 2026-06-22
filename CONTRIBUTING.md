# Contributing

Thank you for improving Auto Docs.

## Development Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python run.py
```

## Build Windows App

```powershell
powershell -ExecutionPolicy Bypass -File packaging\build_windows.ps1
```

## Pull Request Guidelines

- Keep runtime data out of commits: `data/`, `exports/`, `build/`, and `dist/`.
- Run `python -m compileall autodocs scripts run.py` before opening a PR.
- Keep UI text friendly for non-technical Vietnamese office users.
- Prefer small, focused changes.

