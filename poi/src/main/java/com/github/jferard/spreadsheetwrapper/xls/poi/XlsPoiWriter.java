/*******************************************************************************
 *     SpreadsheetWrapper - An abstraction layer over some APIs for Excel or Calc
 *     Copyright (C) 2015  J. Férard
 *
 *     This program is free software: you can redistribute it and/or modify
 *     it under the terms of the GNU General Public License as published by
 *     the Free Software Foundation, either version 3 of the License, or
 *     (at your option) any later version.
 *
 *     This program is distributed in the hope that it will be useful,
 *     but WITHOUT ANY WARRANTY; without even the implied warranty of
 *     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *     GNU General Public License for more details.
 *
 *     You should have received a copy of the GNU General Public License
 *     along with this program.  If not, see <http://www.gnu.org/licenses/>.
 *******************************************************************************/
package com.github.jferard.spreadsheetwrapper.xls.poi;

import java.util.Date;
import java.util.List;

import org.apache.poi.ss.formula.FormulaParseException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.xls.XlsConstants;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 */
class XlsPoiWriter extends AbstractSpreadsheetWriter
		implements SpreadsheetWriter {
	/** current row index, -1 if none */
	private int curR;

	/** current row, null if none */
	private/*@Nullable*/Row curRow;

	/**
	 * the style for cell dates, since Excel does not know the difference
	 * between a date and a number
	 */
	private final/*@Nullable*/CellStyle dateCellStyle;

	/** internal sheet */
	private final Sheet sheet;

	/** helper for style */
	private final XlsPoiStyleHelper styleHelper;

	private static boolean nullCellOrInvalidType(final /*@Nullable*/ Cell cell,
			final int... validTypes) {
		if (cell != null) {
			final int cellType = cell.getCellType();
			for (final int validType : validTypes) {
				if (validType == cellType)
					return false;
			}
		}
		return true;
	}

	/**
	 * @param delegateStyleHelper
	 */
	XlsPoiWriter(final Sheet sheet, final XlsPoiStyleHelper styleHelper,
			final/*@Nullable*/CellStyle dateCellStyle) {
		this.styleHelper = styleHelper;
		this.sheet = sheet;
		this.dateCellStyle = dateCellStyle;

	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Boolean getBoolean(final int r, final int c) {
		final Cell cell = this.getPOICell(r, c);
		if (XlsPoiWriter.nullCellOrInvalidType(cell, Cell.CELL_TYPE_BOOLEAN))
			return null;

		return cell.getBooleanCellValue();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Object getCellContent(final int rowIndex,
			final int colIndex) {
		final Cell cell = this.getPOICell(rowIndex, colIndex);
		if (cell == null)
			return null;

		Object result;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			result = null;
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			result = Boolean.valueOf(cell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_ERROR:
			result = "#Err";
			break;
		case Cell.CELL_TYPE_FORMULA:
			result = cell.getCellFormula();
			break;
		case Cell.CELL_TYPE_NUMERIC:
			if (this.isDate(cell)) {
				result = cell.getDateCellValue();
			} else {
				final double value = cell.getNumericCellValue();
				if (value == Math.rint(value))
					result = Integer.valueOf((int) value);
				else
					result = value;
			}
			break;
		case Cell.CELL_TYPE_STRING:
			result = cell.getStringCellValue();
			break;
		default:
			throw new IllegalArgumentException(String
					.format("Unknown type of cell %d", cell.getCellType()));
		}
		return result;
	}

	/** {@inheritDoc} */
	@Override
	public int getCellCount(final int r) {
		if (r < 0)
			throw new IllegalArgumentException();
		
		final int ret;
		if (r < this.getRowCount())
			ret = this.getCellCountUnsafe(r);
		else
			ret = 0;
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Date getDate(final int r, final int c) {
		final Cell cell = this.getPOICell(r, c);
		if (cell == null || !this.isDate(cell))
			return null;

		return cell.getDateCellValue();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Double getDouble(final int r, final int c) {
		final Cell cell = this.getPOICell(r, c);
		if (cell == null || !this.isDouble(cell))
			return null;

		return cell.getNumericCellValue();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getFormula(final int r, final int c) {
		final Cell cell = this.getPOICell(r, c);
		if (XlsPoiWriter.nullCellOrInvalidType(cell, Cell.CELL_TYPE_FORMULA))
			return null;

		return cell.getCellFormula();
	}

	/** {@inheritDoc} */
	@Override
	public String getName() {
		return this.sheet.getSheetName();
	}

	/** {@inheritDoc} */
	@Override
	public int getRowCount() {
		int rowCount = this.sheet.getLastRowNum() + 1; // 0-based
		if (rowCount == 1 && this.getCellCountUnsafe(0) == 0)
			rowCount = 0;
		return rowCount;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/WrapperCellStyle getStyle(final int r, final int c) {
		final Cell poiCell = this.getPOICell(r, c);
		return this.styleHelper.getStyle(this.sheet.getWorkbook(), poiCell);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getStyleName(final int r, final int c) {
		final Cell cell = this.getPOICell(r, c);
		return this.styleHelper.getStyleName(cell);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getText(final int r, final int c) {
		final Cell cell = this.getPOICell(r, c);
		if (XlsPoiWriter.nullCellOrInvalidType(cell, Cell.CELL_TYPE_STRING))
			return null;
		
		return cell.getStringCellValue();
	}

	/** {@inheritDoc} */
	@Override
	public boolean insertCol(final int c) {
		if (c < 0)
			throw new IllegalArgumentException();

		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public boolean insertRow(final int r) {
		if (r < 0)
			throw new IllegalArgumentException();

		final int rowCount = this.getRowCount();
		if (r < rowCount) {
			this.sheet.shiftRows(r, rowCount, 1);
			this.sheet.createRow(r);
			return true;
		} else
			return false;
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> removeCol(final int c) {
		if (c < 0)
			throw new IllegalArgumentException();
		else
			throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> removeRow(final int r) {
		if (r < 0)
			throw new IllegalArgumentException();
		
		if (r < this.getRowCount()) {
			final List</*@Nullable*/Object> ret = this.getRowContents(r);
			final Row row = this.sheet.getRow(r);
			this.sheet.removeRow(row);
			this.sheet.shiftRows(r, this.getRowCount(), 1);
			return ret;
		} else
			return null;
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public Boolean setBoolean(final int r, final int c, final Boolean value) {
		final Cell cell = this.getOrCreatePOICell(r, c);
		cell.setCellValue(value);
		return value;
	}

	/**
	 * @return
	 */
	@Override
	public Date setDate(final int r, final int c, final Date date) {
		final Cell cell = this.getOrCreatePOICell(r, c);
		cell.setCellValue(date);
		if (this.dateCellStyle != null)
			cell.setCellStyle(this.dateCellStyle);
		return date;
	}

	/**
	 * @return
	 */
	@Override
	public Double setDouble(final int r, final int c, final Number value) {
		final Cell cell = this.getOrCreatePOICell(r, c);
		final double retValue = value.doubleValue();
		cell.setCellValue(retValue);
		return retValue;
	}

	/**
	 * @return
	 */
	@Override
	public String setFormula(final int r, final int c, final String formula) {
		final Cell cell = this.getOrCreatePOICell(r, c);
		try {
			cell.setCellFormula(formula);
		} catch (final FormulaParseException e) {
			throw new IllegalArgumentException(e);
		}
		return formula;
	}

	/**
	 * @return
	 */
	@Override
	public Integer setInteger(final int r, final int c, final Number value) {
		final Cell cell = this.getOrCreatePOICell(r, c);
		final int retValue = value.intValue();
		cell.setCellValue(retValue);
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyle(final int r, final int c,
			final WrapperCellStyle wrapperCellStyle) {
		final Cell poiCell = this.getOrCreatePOICell(r, c);
		return this.styleHelper.setStyle(this.sheet.getWorkbook(), poiCell,
				wrapperCellStyle);
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c,
			final String styleName) {
		final Cell cell = this.getOrCreatePOICell(r, c);
		return this.styleHelper.setSyleName(this.sheet.getWorkbook(), cell,
				styleName);
	}

	/** */
	@Override
	public String setText(final int r, final int c, final String text) {
		final Cell cell = this.getOrCreatePOICell(r, c);
		cell.setCellValue(text);
		return text;
	}

	private int getCellCountUnsafe(final int r) {
		final Row row = this.sheet.getRow(r);
		final short count;
		if (row == null)
			count = 0;
		else
			count = row.getLastCellNum(); // 1-based

		return count;
	}

	/**
	 * Simple optimization hidden inside a method.
	 *
	 * @param r
	 *            the row index
	 * @param c
	 *            the column index
	 * @return the cell
	 */
	private/*@Nullable*/Cell getPOICell(final int r, final int c) {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();
		if (r >= this.getRowCount() || c >= this.getCellCount(r))
			return null;

		if (r != this.curR || this.curRow == null) {
			this.curRow = this.sheet.getRow(r);
			assert this.curRow != null; // else this.curRow =
			// this.sheet.createRow(r);
			this.curR = r;
		}
		Cell cell = this.curRow.getCell(c);
		if (cell == null)
			cell = this.curRow.createCell(c);
		return cell;
	}

	/**
	 * Simple optimization hidden inside a method.
	 *
	 * @param r
	 *            the row index
	 * @param c
	 *            the column index
	 * @return the cell
	 */
	protected Cell getOrCreatePOICell(final int r, final int c) {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();
		if (r >= XlsConstants.MAX_ROWS_PER_SHEET
				|| c >= XlsConstants.MAX_COLUMNS)
			throw new UnsupportedOperationException();

		if (r != this.curR || this.curRow == null) {
			this.curRow = this.sheet.getRow(r);
			if (this.curRow == null)
				this.curRow = this.sheet.createRow(r);
			this.curR = r;
		}
		Cell cell = this.curRow.getCell(c);
		if (cell == null)
			cell = this.curRow.createCell(c);
		return cell;
	}

	/**
	 * @param cell
	 *            the cell
	 * @return true if it's a (fake) date cell
	 */
	protected boolean isDate(final Cell cell) {
		return cell.getCellType() == Cell.CELL_TYPE_NUMERIC
				&& DateUtil.isCellDateFormatted(cell);
	}

	/**
	 * @param cell
	 *            the cell
	 * @return true if it's not a (fake) date cell
	 */
	protected boolean isDouble(final Cell cell) {
		return cell.getCellType() == Cell.CELL_TYPE_NUMERIC
				&& !DateUtil.isCellDateFormatted(cell);
	}
}