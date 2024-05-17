
helps to create cutom excel sheet from LR analysis default SUMMARY REPORT EXCEL ouput:

==============================
==============================
==============================
==============================
==============================
==============================
MainModule.vba: Main entry point containing the Sub aaaaaaaaaRunMeForAnalysisSheetGeneration() which orchestrates the execution of other subroutines.

RenameSheetModule.vba: Contains the subroutine RenameSheetIfInvalid() for renaming sheets if invalid.

CopyTransactionModule.vba: Contains the subroutine CopyTransactionSummary() for copying transaction summaries.

SortAndMoveModule.vba: Contains the subroutine SortAndMoveData() for sorting and moving data.

CopyHeadingModule.vba: Contains the subroutine CopyHeadingRowToSplitData() for copying heading rows to the split data sheet.

DeleteColumnsSLAModule.vba: Contains the subroutine DeleteColumnsStartingWithSLAStatus() for deleting columns starting with "SLA Status".

DeleteColumnsStopModule.vba: Contains the subroutine DeleteColumnsStartingWithStop() for deleting columns starting with "Stop".

FindColorEmptyRowsModule.vba: Contains the subroutine FindAndColorEmptyRows() for finding and coloring empty rows.

AutoFitDataModule.vba: Contains the subroutine AutoFitDatainSplitDataSheet() for autofitting data in the split data sheet.
