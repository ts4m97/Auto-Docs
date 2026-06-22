# Release Guide

This project builds a Windows desktop app with PyInstaller.

## Build

```powershell
powershell -ExecutionPolicy Bypass -File packaging\build_windows.ps1
```

Output:

```text
dist\AutoDocs\AutoDocs.exe
```

## Suggested Release Checklist

1. Update `CHANGELOG.md`.
2. Update version numbers in:
   - `pyproject.toml`
   - `packaging/windows_version_info.txt`
3. Run:
   ```powershell
   python -m compileall autodocs scripts run.py
   ```
4. Build the app.
5. Launch `dist\AutoDocs\AutoDocs.exe`.
6. Test:
   - Import template
   - Manual export DOCX
   - Excel import preview
   - Export history
   - Batch print queue UI
7. Zip `dist\AutoDocs`.
8. Attach the zip to a GitHub Release.

