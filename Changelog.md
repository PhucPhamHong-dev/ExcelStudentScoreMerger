All notable changes to this project are documented in this file.

[1.1.0] - 2025-06-15

Added

Switched all file operations to use a single TemporaryDirectory to eliminate persistent temp/output folders.

Filename deduplication: output files now append (2), (3), etc. when multiple templates share the same exam code.

ZIP archive now includes only generated _filled.xlsx files; raw and template files are excluded.

Dynamic download naming: for a single code, the ZIP is named <code>_filled.zip; otherwise FE_Merge.zip.

Fixed

Bug where FormData used the wrong variable name (form → now fd).

Corrected front-end fetch integration to use fd consistently.

TemplateNotFound error fixed by configuring Flask template_folder correctly.

[1.0.0] - Initial Release

Basic Flask app to merge raw exam scores into multiple Excel templates.

Front-end with TailwindCSS, AlpineJS, and SheetJS for validation.

Server-side merge logic using pandas and openpyxl.

Outputs zipped as FE_Merge.zip containing all filled result files.

