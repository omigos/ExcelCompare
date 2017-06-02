package com.ka.spreadsheet.diff;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Workbook;

import javax.annotation.Nullable;

/**
 * All indexes are zero based
 */
public interface ISpreadSheet {

  Iterator<ISheet> getSheetIterator();

  IFont getFont(short index);

  @Nullable
  Boolean hasMacro();
}


interface ISheet {

  String getName();

  int getSheetIndex();

  Iterator<IRow> getRowIterator();
}


interface IRow {

  int getRowIndex();

  Iterator<ICell> getCellIterator();
}


interface ICell {

  int getRowIndex();

  int getColumnIndex();

  CellValue getValue();

  ICellStyle getCellStyle();
}


interface ICellStyle {
  int getAlignment();

  short getBorderBottom();

  short getBorderLeft();

  short getBorderRight();

  short getBorderTop();

  short getBottomBorderColor();

  short getLeftBorderColor();

  short getTopBorderColor();

  short getRightBorderColor();

  String getDataFormatString();

  short getFillBackgroundColor();

  Color getFillBackgroundColorColor();

  short getFillForegroundColor();

  Color getFillForegroundColorColor();

  short getFillPattern();

  boolean getHidden();

  short getIndention();

  short getVerticalAlignment();

  boolean getWrapText();

  short getRotation();

  boolean getLocked();

  short getFontIndex();
}


interface IFont {

  short getBoldweight();

  short getColor();

  short getFontHeight();

  String getFontName();
}