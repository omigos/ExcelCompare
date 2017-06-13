package com.ka.spreadsheet.diff;

import static com.ka.spreadsheet.diff.Flags.WORKBOOK1;
import static com.ka.spreadsheet.diff.Flags.WORKBOOK2;

import java.io.File;
import java.util.Iterator;

import org.apache.poi.hssf.util.PaneInformation;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.odftoolkit.simple.SpreadsheetDocument;


public class SpreadSheetDiffer {

  public static void main(String[] args) {
    int ret = doDiff(args);
    System.exit(ret);
  }

  public static int doDiff(String[] args) {
    int ret = -1;
    try {
      if (Flags.parseFlags(args)) {
        ret = doDiff(new StdoutSpreadSheetDiffCallback());
      }
    } catch (Exception e) {
      if (Flags.DEBUG) {
        e.printStackTrace(System.err);
      } else {
        System.err.println("Diff failed: " + e.getMessage());
      }
    }
    return ret;
  }

  public static int doDiff(SpreadSheetDiffCallback diffCallback) throws Exception {
    if (!verifyFile(WORKBOOK1) || !verifyFile(WORKBOOK2)) {
      return -1;
    }

    ISpreadSheet ss1 = isDevNull(WORKBOOK1) ? emptySpreadSheet() : loadSpreadSheet(WORKBOOK1);
    ISpreadSheet ss2 = isDevNull(WORKBOOK2) ? emptySpreadSheet() : loadSpreadSheet(WORKBOOK2);

    ISpreadSheetIterator ssi1 = isDevNull(WORKBOOK1) ?
        emptySpreadSheetIterator() : new SpreadSheetIterator(ss1, Flags.WORKBOOK_IGNORES1);
    ISpreadSheetIterator ssi2 = isDevNull(WORKBOOK2) ?
        emptySpreadSheetIterator() : new SpreadSheetIterator(ss2, Flags.WORKBOOK_IGNORES2);

    boolean isDiff = false;
    CellPos c1 = null, c2 = null;
    while (true) {
      if ((c1 == null) && ssi1.hasNext())
        c1 = ssi1.next();
      if ((c2 == null) && ssi2.hasNext())
        c2 = ssi2.next();

      if ((c1 != null) && (c2 != null)) {
        int c = c1.compareCellPositions(c2);
        if (c == 0) {
          if (!c1.getCellValue().compare(c2.getCellValue())) {
            isDiff = true;
            diffCallback.reportDiffCell(c1, c2);
          } else {
            try {
              styleCompare(c1, c2, ss1, ss2);
            } catch (Exception e) {
              isDiff = true;
              diffCallback.reportStyleDiff(e.getMessage(), c1, c2);
            }
          }
          c1 = c2 = null;
        } else if (c < 0) {
          isDiff = true;
          diffCallback.reportExtraCell(true, c1);
          c1 = null;
        } else {
          isDiff = true;
          diffCallback.reportExtraCell(false, c2);
          c2 = null;
        }
      } else {
        break;
      }
    }
    if ((c1 != null) && (c2 == null)) {
      do {
        isDiff = true;
        diffCallback.reportExtraCell(true, c1);
        c1 = ssi1.hasNext() ? ssi1.next() : null;
      } while (c1 != null);
    } else if ((c1 == null) && (c2 != null)) {
      do {
        isDiff = true;
        diffCallback.reportExtraCell(false, c2);
        c2 = ssi2.hasNext() ? ssi2.next() : null;
      } while (c2 != null);
    }
    if ((c1 != null) || (c2 != null)) {
      throw new IllegalStateException("Something wrong");
    }


    if (ss1 instanceof SpreadSheetExcel && ss2 instanceof SpreadSheetExcel) {
      SpreadSheetExcel sse1 = (SpreadSheetExcel) ss1;
      SpreadSheetExcel sse2 = (SpreadSheetExcel) ss2;

      int i = 0;
      while (true) {
        if (i >= sse1.getWorkbook().getNumberOfSheets() ||
                i >= sse2.getWorkbook().getNumberOfSheets())
          break;

        Sheet s1 = sse1.getWorkbook().getSheetAt(i);
        Sheet s2 = sse2.getWorkbook().getSheetAt(i);

        if (s1 == null || s2 == null ||
                !(s1 instanceof XSSFSheet) ||
                !(s2 instanceof XSSFSheet) ||
                !(sse1.getWorkbook() instanceof XSSFWorkbook) ||
                !(sse2.getWorkbook() instanceof XSSFWorkbook))
          break;

        XSSFSheet xs1 = (XSSFSheet) s1;
        XSSFSheet xs2 = (XSSFSheet) s2;
        XSSFWorkbook w1 = (XSSFWorkbook) sse1.getWorkbook();
        XSSFWorkbook w2 = (XSSFWorkbook) sse2.getWorkbook();

        try {
          compareColumnWidth(xs1, xs2);
        } catch (Exception e) {
          isDiff = true;
          diffCallback.reportSimpleDiff(e.getMessage(), xs1, xs2);
        }
        try {
          compareFreezePane(xs1, xs2);
        } catch (Exception e) {
          isDiff = true;
          diffCallback.reportSimpleDiff(e.getMessage(), xs1, xs2);
        }
        try {
          compareMergedRegion(xs1, xs2);
        } catch (Exception e) {
          isDiff = true;
          diffCallback.reportSimpleDiff(e.getMessage(), xs1, xs2);
        }
        try {
          compareRowGroup(xs1, xs2);
        } catch (Exception e) {
          isDiff = true;
          diffCallback.reportSimpleDiff(e.getMessage(), xs1, xs2);
        }
        try {
          compareSheetName(xs1, xs2);
        } catch (Exception e) {
          isDiff = true;
          diffCallback.reportSimpleDiff(e.getMessage(), xs1, xs2);
        }
        i++;
      }
    }

    Boolean hasMacro1 = ss1.hasMacro();
    Boolean hasMacro2 = ss2.hasMacro();
    if ((hasMacro1 != null) && (hasMacro2 != null) && (hasMacro1 != hasMacro2)) {
      isDiff = true;
      diffCallback.reportMacroOnlyIn(hasMacro1);
    }

    diffCallback.reportWorkbooksDiffer(isDiff, WORKBOOK1, WORKBOOK2);

    return isDiff ? 1 : 0;
  }

  private static void compareRowGroup(XSSFSheet xs1, XSSFSheet xs2) {
    // RowGroup
    Iterator<Row> ri1 = xs1.iterator();
    Iterator<Row> ri2 = xs2.iterator();
    while (ri1.hasNext() && ri2.hasNext()) {
      Row r1 = ri1.next();
      Row r2 = ri2.next();

      verifyStyle(r1.getOutlineLevel(), r2.getOutlineLevel(), "RowGroup.OutlineLevel for row " + r1.getRowNum());
    }
  }

  private static void compareColumnWidth(XSSFSheet xs1, XSSFSheet xs2) {
    // ColumnWidth
    Iterator<Row> rowIterator = xs1.iterator();
    short maxColumn = 0;
    while (rowIterator.hasNext()) {
      Row row = rowIterator.next();
      if (row.getLastCellNum() > maxColumn)
        maxColumn = row.getLastCellNum();
    }
    for (int ii = 0; ii < maxColumn; ii++) {
      verifyStyle(xs1.getColumnWidth(ii), xs2.getColumnWidth(ii), "ColumnWidth for column " + ii);
    }
  }

  private static void compareFreezePane(XSSFSheet xs1, XSSFSheet xs2) {
    // FreezePane
    PaneInformation pi1 = xs1.getPaneInformation();
    PaneInformation pi2 = xs2.getPaneInformation();
    verifyStyle(pi1 == null ? "null" : "not null", pi2 == null ? "null" : "not null", "FreezePane");
    if (pi1 != null && pi2 != null) {
      verifyStyle(pi1.isFreezePane(), pi2.isFreezePane(), "FreezePane.isFreezePane");
      verifyStyle(pi1.getActivePane(), pi2.getActivePane(), "FreezePane.getActivePane");
      verifyStyle(pi1.getHorizontalSplitPosition(), pi2.getHorizontalSplitPosition(), "FreezePane.getHorizontalSplitPosition");
      verifyStyle(pi1.getHorizontalSplitTopRow(), pi2.getHorizontalSplitTopRow(), "FreezePane.getHorizontalSplitTopRow");
      verifyStyle(pi1.getVerticalSplitLeftColumn(), pi2.getVerticalSplitLeftColumn(), "FreezePane.getVerticalSplitLeftColumn");
      verifyStyle(pi1.getVerticalSplitPosition(), pi2.getVerticalSplitPosition(), "FreezePane.getVerticalSplitPosition");
    }
  }

  private static void compareSheetName(XSSFSheet xs1, XSSFSheet xs2) {
    // SheetName
    verifyStyle(xs1.getSheetName(), xs2.getSheetName(), "sheetName");
  }

  private static void compareMergedRegion(XSSFSheet xs1, XSSFSheet xs2) {
    // MergedRegion
    if (xs1.getNumMergedRegions() != xs2.getNumMergedRegions()) {
      verifyStyle(xs1.getNumMergedRegions(), xs2.getNumMergedRegions(), "MergedRegion.getNumMergedRegions");
    } else {
      for (int ii = 0; ii < xs1.getNumMergedRegions(); ii++) {
        CellRangeAddress mergedRegion1 = xs1.getMergedRegion(ii);
        CellRangeAddress mergedRegion2 = xs2.getMergedRegion(ii);

        verifyStyle(mergedRegion1.getFirstColumn(), mergedRegion2.getFirstColumn(), "MergedRegion.getFirstColumn");
        verifyStyle(mergedRegion1.getFirstRow(), mergedRegion2.getFirstRow(), "MergedRegion.getFirstRow");
        verifyStyle(mergedRegion1.getLastColumn(), mergedRegion2.getLastColumn(), "MergedRegion.getLastColumn");
        verifyStyle(mergedRegion1.getLastRow(), mergedRegion2.getLastRow(), "MergedRegion.getLastRow");
        verifyStyle(mergedRegion1.getNumberOfCells(), mergedRegion2.getNumberOfCells(), "MergedRegion.getNumberOfCells");
        verifyStyle(mergedRegion1.isFullColumnRange(), mergedRegion2.isFullColumnRange(), "MergedRegion.isFullColumnRange");
        verifyStyle(mergedRegion1.isFullRowRange(), mergedRegion2.isFullRowRange(), "MergedRegion.isFullRowRange");
      }
    }
  }

  private static boolean isDevNull(File file) {
    return "/dev/null".equals(file.getAbsolutePath())
        || "\\\\.\\NUL".equals(file.getAbsolutePath());
  }

  private static boolean verifyFile(File file) {
    if (isDevNull(file)) {
      return true;
    }
    if (!file.exists()) {
      System.err.println("File: " + file.toString().replace("\\", "/") + " does not exist.");
      return false;
    }
    if (!file.canRead()) {
      System.err.println("File: " + file.toString().replace("\\", "/") + " not readable.");
      return false;
    }
    if (!file.isFile()) {
      System.err.println("File: " + file.toString().replace("\\", "/") + " is not a file.");
      return false;
    }
    return true;
  }

  private static ISpreadSheet loadSpreadSheet(File file) throws Exception {
    // assume file is excel by default
    Exception excelReadException = null;
    try {
      Workbook workbook = WorkbookFactory.create(file);
      return new SpreadSheetExcel(workbook);
    } catch (Exception e) {
      excelReadException = e;
    }
    Exception odfReadException = null;
    try {
      SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.loadDocument(file);
      return new SpreadSheetOdf(spreadsheetDocument);
    } catch (Exception e) {
      odfReadException = e;
    }
    if (file.getName().matches(".*\\.ods.*")) {
      throw new RuntimeException("Failed to read as ods file: " + file.toString().replace("\\", "/"), odfReadException);
    } else {
      throw new RuntimeException("Failed to read as excel file: " + file.toString().replace("\\", "/"), excelReadException);
    }
  }

  private static ISpreadSheet emptySpreadSheet() {
    return new ISpreadSheet() {
      @Override
      public Boolean hasMacro() {
        return false;
      }
      @Override
      public Iterator<ISheet> getSheetIterator() {
        return new Iterator<ISheet>() {
          @Override
          public boolean hasNext() {
            return false;
          }
          @Override
          public ISheet next() {
            throw new IllegalStateException();
          }
          @Override
          public void remove() {
            throw new IllegalStateException();
          }
        };
      }

      @Override
      public IFont getFont(short index) {
        return null;
      }
    };
  }

  private static ISpreadSheetIterator emptySpreadSheetIterator() {
    return new ISpreadSheetIterator() {
      @Override
      public boolean hasNext() {
        return false;
      }
      @Override
      public CellPos next() {
        throw new IllegalStateException();
      }
    };
  }

  private static void styleCompare(CellPos c1, CellPos c2, ISpreadSheet ss1, ISpreadSheet ss2) {
    ICellStyle s1 = c1.getCell().getCellStyle();
    ICellStyle s2 = c2.getCell().getCellStyle();

    try {
      verifyStyle(s1.getLocked(), s2.getLocked(), "locked");
      verifyStyle(s1.getAlignment(), s2.getAlignment(), "alignment");
      verifyStyle(s1.getBorderBottom(), s2.getBorderBottom(), "borderBottom");
      verifyStyle(s1.getBorderLeft(), s2.getBorderLeft(), "borderLeft");
      verifyStyle(s1.getBorderRight(), s2.getBorderRight(), "borderRight");
      verifyStyle(s1.getBorderTop(), s2.getBorderTop(), "borderTop");
      verifyStyle(s1.getWrapText(), s2.getWrapText(), "wrapText");
      verifyStyle(s1.getVerticalAlignment(), s2.getVerticalAlignment(), "verticalAlignment");
      verifyStyle(s1.getTopBorderColor(), s2.getTopBorderColor(), "topBorderColor");
      verifyStyle(s1.getRotation(), s2.getRotation(), "rotation");
      verifyStyle(s1.getRightBorderColor(), s2.getRightBorderColor(), "rightBorderColor");
      verifyStyle(s1.getLeftBorderColor(), s2.getLeftBorderColor(), "leftBorderColor");
      verifyStyle(s1.getIndention(), s2.getIndention(), "indention");
      verifyStyle(s1.getHidden(), s2.getHidden(), "hidden");
      verifyStyle(s1.getFillPattern(), s2.getFillPattern(), "fillPattern");
      verifyStyle(s1.getFillForegroundColorColor(), s2.getFillForegroundColorColor(), "fillForegroundColorColor");
      verifyStyle(s1.getFillForegroundColor(), s2.getFillForegroundColor(), "fillForegroundColor");
      verifyStyle(s1.getDataFormatString(), s2.getDataFormatString(), "dataFormatString");
      verifyStyle(s1.getBottomBorderColor(), s2.getBottomBorderColor(), "bottomBordercolor");
      verifyStyle(s1.getFillBackgroundColor(), s2.getFillBackgroundColor(), "fillBackgroundColor");
      verifyStyle(s1.getFillBackgroundColorColor(), s2.getFillBackgroundColorColor(), "fillBackgroundColorColor");

    } catch (IllegalStateException e) {
      throw new IllegalStateException("Styles of Cell " + c1.getCellPosition() + " does not match " + c2.getCellPosition() + " (" + e.getMessage() + ")");
    }

    if (c1.getCellValue().toString().trim().equals("") && c2.getCellValue().toString().trim().equals(""))
      return;

    IFont f1;
    try {
      f1 = ss1.getFont(c1.getCell().getCellStyle().getFontIndex());
    } catch (Exception e) {
      throw new IllegalStateException("failed to load font 1 #" + c1.getCell().getCellStyle().getFontIndex());
    }

    IFont f2;
    try {
      f2 = ss2.getFont(c2.getCell().getCellStyle().getFontIndex());
    } catch (Exception e) {
      throw new IllegalStateException("failed to load font 2 #" + c1.getCell().getCellStyle().getFontIndex());
    }

    try {
      if (f1 != null && f2 != null) {
        verifyStyle(f1.getBoldweight(), f2.getBoldweight(), "bold");
        verifyStyle(f1.getColor(), f2.getColor(), "color");
        verifyStyle(f1.getFontHeight(), f2.getFontHeight(), "fontHeight");
        verifyStyle(f1.getFontName(), f2.getFontName(), "fontName");
      }
    } catch (IllegalStateException e) {
      throw new IllegalStateException("Styles of Cell " + c1.getCellPosition() + " does not match " + c2.getCellPosition() + " (" + e.getMessage() + ") for content '" + c1.getCellValue().toString() + "'");
    }
  }

  private static void verifyStyle(Object o1, Object o2, String description) {
    if (o1 == null && o2 == null)
      return;

    if (o1 == null || o2 == null) {
      throw new IllegalStateException(description + " do not match: " + o1 + " != " + o2);
    }
    if (!o1.equals(o2)) {
      throw new IllegalStateException(description + " do not match: " + o1 + " != " + o2);
    }
  }

}
