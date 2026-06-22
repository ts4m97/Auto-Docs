# Development Notes

## Architecture

Auto Docs is intentionally split into small modules:

- `autodocs/main.py` handles the PySide6 interface.
- `autodocs/document_service.py` handles Word, Excel, PDF, and image conversion.
- `autodocs/storage.py` handles SQLite persistence for templates and export history.
- `autodocs/print_service.py` handles Windows printing.

## Runtime Data

Runtime data is local and ignored by Git:

- `data/`
- `exports/`

When running from the PyInstaller build, these folders are created next to
`AutoDocs.exe`.

## Regenerating Assets

```powershell
python scripts\create_app_icon.py
python scripts\create_project_branding.py
```

## Regenerating Test Templates

```powershell
python scripts\create_test_templates.py
```

