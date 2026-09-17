# Changelog

## 3.1.0 - 2026-09-17

### Regression fixes
- Restored the classic undo workflow and added `Ctrl+Z` history for add/update/delete/duplicate/import operations.
- Restored a dedicated PDF/table title instead of overloading the employee field.
- Restored the classic quick total-hours dialog.
- Restored PDF-to-JPG conversion while keeping the newer PNG export.
- Kept existing v3.0 state files compatible through automatic schema migration.

### Improved
- Added localized PDF headers for Polish, Norwegian and English.
- Added the custom application icon to the README and Windows GUI runtime assets.
- Added an offline GUI smoke mode, Windows GUI CI and packaged EXE GUI smoke testing.
- Updated icon tooling to generate both ICO and PNG artwork from the same design.

## 3.0.0 - 2026-09-17

### Changed
- Rebuilt the project from three language/version scripts into one maintainable package.
- Added automatic system-language selection for Polish, Norwegian and English with English fallback.
- Added a modern dark-blue responsive desktop interface.
- Moved runtime data to the user profile instead of the repository/application folder.
- Replaced separate language executables with one multilingual Windows application.

### Added
- Live field normalization and one-click repair for common date/time/client formatting problems.
- Overnight-shift calculation and configurable break length.
- Row editing, deletion and duplication.
- Autosave, CSV import/export and landscape PDF export.
- Optional PDF-to-PNG conversion.
- Client summary totals.
- Custom Worker Time List application icon.
- Python 3.10-3.14 CI.
- Automated Windows EXE, portable ZIP and SHA256 release pipeline.

### Removed
- `worker v1.py`
- `worker v2.py`
- `worker_v2_EN.py`
- dependency on duplicate per-language application code.
