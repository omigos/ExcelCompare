package com.ka.spreadsheet.diff;

import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;

public class SpreadSheetExcel implements ISpreadSheet {

	private Workbook workbook;
	
	public SpreadSheetExcel(Workbook workbook) {
		this.workbook = workbook;
	}

    public IFont getFont(short index) {
        return new FontExcel(workbook.getFontAt(index));
    }
	
	@Override
	public Iterator<ISheet> getSheetIterator() {
		return new Iterator<ISheet>() {
			
			private int currSheetIdx = 0;

			@Override
			public boolean hasNext() {
				return currSheetIdx < workbook.getNumberOfSheets();
			}

			@Override
			public ISheet next() {
				Sheet sheet = workbook.getSheetAt(currSheetIdx);
				SheetExcel sheetExcel = new SheetExcel(sheet, currSheetIdx);
				currSheetIdx++;
				return sheetExcel;
			}

			@Override
			public void remove() {
				throw new UnsupportedOperationException();
			}
		};
	}
}

class SheetExcel implements ISheet {

	private Sheet sheet;
	private int sheetIdx;
	
	public SheetExcel(Sheet sheet, int sheetIdx) {
		this.sheet = sheet;
		this.sheetIdx = sheetIdx;
	}
	
	@Override
	public String getName() {
		return sheet.getSheetName();
	}
	
	@Override
	public int getSheetIndex() {
		return sheetIdx;
	}

	@Override
	public Iterator<IRow> getRowIterator() {
		final Iterator<Row> rowIterator = sheet.rowIterator();
		return new Iterator<IRow>() {
			
			@Override
			public boolean hasNext() {
				return rowIterator.hasNext();
			}

			@Override
			public IRow next() {
				return new RowExcel(rowIterator.next());
			}

			@Override
			public void remove() {
				throw new UnsupportedOperationException();
			}
		};
	}
}

class RowExcel implements IRow {
	
	private Row row;
	
	public RowExcel(Row row) {
		this.row = row;
	}

	@Override
	public int getRowIndex() {
		return row.getRowNum();
	}

	@Override
	public Iterator<ICell> getCellIterator() {
		final Iterator<Cell> cellIterator = row.cellIterator();
		return new Iterator<ICell>() {
			
			@Override
			public boolean hasNext() {
				return cellIterator.hasNext();
			}

			@Override
			public ICell next() {
				return new CellExcel(cellIterator.next());
			}

			@Override
			public void remove() {
				throw new UnsupportedOperationException();
			}
		};
	}
}

class CellExcel implements ICell {

	private Cell cell;
	
	public CellExcel(Cell cell) {
		this.cell = cell;
	}

	@Override
	public int getRowIndex() {
		return cell.getRowIndex();
	}

	@Override
	public int getColumnIndex() {
		return cell.getColumnIndex();
	}

    @Override
    public ICellStyle getCellStyle() {
        return new CellStyleExcel(cell.getCellStyle());
    }

    @Override
	public String getStringValue() {
		int cellType = cell.getCellType();
		switch (cellType) {
		case Cell.CELL_TYPE_NUMERIC:
			return String.valueOf(cell.getNumericCellValue());
		case Cell.CELL_TYPE_BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		case Cell.CELL_TYPE_BLANK:
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		case Cell.CELL_TYPE_FORMULA:
			return cell.getCellFormula();
		case Cell.CELL_TYPE_ERROR:
			return String.valueOf(cell.getErrorCellValue());
		}
		return cell.getStringCellValue();
	}
}

class CellStyleExcel implements ICellStyle {
    private CellStyle cellStyle;

    public CellStyleExcel(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    @Override
    public int getAlignment() {
        return cellStyle.getAlignment();
    }

    @Override
    public short getBorderBottom() {
        return cellStyle.getBorderBottom();
    }

    @Override
    public short getBorderLeft() {

        return cellStyle.getBorderLeft();
    }

    @Override
    public short getBorderRight() {
        return cellStyle.getBorderRight();
    }

    @Override
    public short getBorderTop() {
        return cellStyle.getBorderTop();
    }

    @Override
    public short getBottomBorderColor() {
        return cellStyle.getBottomBorderColor();
    }

    @Override
    public short getLeftBorderColor() {
        return cellStyle.getLeftBorderColor();
    }

    @Override
    public short getTopBorderColor() {
        return cellStyle.getTopBorderColor();
    }

    @Override
    public short getRightBorderColor() {
        return cellStyle.getRightBorderColor();
    }

    @Override
    public String getDataFormatString() {
        return cellStyle.getDataFormatString();
    }

    @Override
    public short getFillBackgroundColor() {
        return cellStyle.getFillBackgroundColor();
    }

    @Override
    public Color getFillBackgroundColorColor() {
        return cellStyle.getFillBackgroundColorColor();
    }

    @Override
    public short getFillForegroundColor() {
        return cellStyle.getFillForegroundColor();
    }

    @Override
    public Color getFillForegroundColorColor() {
        return cellStyle.getFillForegroundColorColor();
    }

    @Override
    public short getFillPattern() {
        return cellStyle.getFillPattern();
    }

    @Override
    public boolean getHidden() {
        return cellStyle.getHidden();
    }

    @Override
    public short getIndention() {
        return cellStyle.getIndention();
    }

    @Override
    public short getVerticalAlignment() {
        return cellStyle.getVerticalAlignment();
    }

    @Override
    public boolean getWrapText() {
        return cellStyle.getWrapText();
    }

    @Override
    public short getRotation() {
        return cellStyle.getRotation();
    }

    @Override
    public boolean getLocked() {
        return cellStyle.getLocked();
    }

    public short getFontIndex() {
        return cellStyle.getFontIndex();
    }
}

class FontExcel implements IFont {
    private Font font;

    public FontExcel(Font font) {
        this.font = font;
    }

    @Override
    public short getBoldweight() {
        return font.getBoldweight();
    }

    @Override
    public short getColor() {
        return font.getColor();
    }

    @Override
    public short getFontHeight() {
        return font.getFontHeight();
    }

    @Override
    public String getFontName() {
        return font.getFontName();
    }
}