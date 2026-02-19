# Changelog

All notable changes to this project will be documented in this file.

## [0.1.4] - 2026-02-18
### Added
- Support for repeated `--source-file` arguments to process specific files directly.
- Helper script `scripts/rerun_failed_historic_events.ps1` to rerun the known failed Historic Events items.

### Fixed
- Automatic deterministic output filename shortening for overlong output paths to avoid Windows/Chromium write failures on long MHT titles.

## [0.1.3] - 2026-02-18
### Fixed
- Hardened Windows path argument parsing against multiline/indent artifacts from shell line continuation mistakes.

## [0.1.2] - 2026-02-18
### Fixed
- Trim accidental leading/trailing whitespace/newlines in `--source-root`, `--output-root`, and `--log-path`.

## [0.1.1] - 2026-02-18
### Fixed
- Added timezone abbreviation mappings in date parsing (including `MST`) to eliminate `UnknownTimezoneWarning` and improve publication-date normalization.

## [0.1.0] - 2026-02-18
### Added
- Initial MHT/MHTML to PDF conversion pipeline.
- Metadata extraction from MHT headers and HTML content.
- Embedded PDF metadata (Info + XMP) and JSON sidecar metadata output.
- Configurable browser selection, output root, log path, recursion behavior, and skip-existing support.
