# excel-to-csv
SAS macro with Python script to import Excel spreadsheet and avoid SAS import errors due to extra newlines and non-UTF-8 characters

Excel is Unicode by default and allows newlines in cells. SAS will generate errors when importing non-UTF-8 Unicode and newlines in cells.

This solution is for when the SAS program must not fail and newlines must be stripped. The non-UTF-8 Unicode chars are replaced with "?".

Make sure to set the paths in the macro.
