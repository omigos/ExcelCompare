package com.ka.spreadsheet.diff;

import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Iterator;

/**
 * All indexes are zero based
 */
public interface ISpreadSheet {

	Iterator<ISheet> getSheetIterator();

    IFont getFont(short index);
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

    ICellStyle getCellStyle();

	String getStringValue();
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