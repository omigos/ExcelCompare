package com.ka.spreadsheet.diff;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.File;

public interface SpreadSheetDiffCallback {

  void reportDiffCell(CellPos c1, CellPos c2);

  void reportExtraCell(boolean inFirstSpreadSheet, CellPos c);

  void reportMacroOnlyIn(boolean inFirstSpreadSheet);

  void reportWorkbooksDiffer(boolean differ, File file1, File file2);

  void reportStyleDiff(String diff, CellPos c1, CellPos c2);

  void reportSimpleDiff(String diff, XSSFSheet xs1, XSSFSheet xs2);
}
