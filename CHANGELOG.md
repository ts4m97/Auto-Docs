# Changelog

## 1.1.0 - 2026-06-25

- Added customer storage with CCCD as the customer ID.
- Added flexible customer fields and placeholder alias mapping.
- Added SQL Server import from `dbo.NguoiLX`, `dbo.NguoiLX_HoSo`, `dbo.KhoaHoc`, and `dbo.DM_DVHC`.
- Added batch export for selected customers from a Word template.
- Improved the customer list selection UI.
- Added `pyodbc` packaging support.

## 1.0.0 - 2026-06-22

- Initial Windows offline desktop app.
- Manage Word `.docx` templates with `[[placeholder]]` fields.
- Fill data manually or from Excel.
- Export DOCX, PDF, and PNG images.
- Preserve placeholder formatting when exporting.
- Track export history and generated files.
- Batch print documents with printer selection and throttling.
- Build Windows executable with embedded icon and version metadata.
