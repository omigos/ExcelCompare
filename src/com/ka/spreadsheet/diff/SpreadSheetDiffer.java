package com.ka.spreadsheet.diff;

import static com.ka.spreadsheet.diff.Flags.WORKBOOK1;
import static com.ka.spreadsheet.diff.Flags.WORKBOOK2;

import java.io.File;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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

    Boolean hasMacro1 = ss1.hasMacro();
    Boolean hasMacro2 = ss2.hasMacro();
    if ((hasMacro1 != null) && (hasMacro2 != null) && (hasMacro1 != hasMacro2)) {
      isDiff = true;
      diffCallback.reportMacroOnlyIn(hasMacro1);
    }

    diffCallback.reportWorkbooksDiffer(isDiff, WORKBOOK1, WORKBOOK2);

    return isDiff ? 1 : 0;
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
