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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetReader;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 */
class XlsPoiReader extends AbstractSpreadsheetReader implements
SpreadsheetReader {
	/** current row index, -1 if none */
	private int curR;
	/** current row, null if none */
	private/*@Nullable*/Row curRow;
	/** *internal* sheet */
	private final Sheet sheet;

	/**
	 * @param sheet
	 *            *internal* sheet
	 */
	XlsPoiReader(final Sheet sheet) {
		super();
		this.sheet = sheet;
		this.curRow = null;
		this.curR = -1;
	}

	/** {@inheritDoc} */
	@Override
	public Boolean getBoolean(final int r, final int c) {
		final Cell cell = this.getPOICell(r, c);
		if (cell.getCellType() != Cell.CELL_TYPE_BOOLEAN)
			throw new IllegalArgumentException(
					"This cell does not contain a boolean");
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
			throw new IllegalArgumentException(String.format(
					"Unknown type of cell %d", cell.getCellType()));
		}
		return result;
	}

	/** {@inheritDoc} */
	@Override
	public int getCellCount(final int r) {
		final Row row = this.sheet.getRow(r);
		final short count;
		if (row == null)
			count = 0;
		else
			count = row.getLastCellNum(); // 1-based

		return count;
	}

	/** {@inheritDoc} */
	@Override
	public Date getDate(final int r, final int c) {
		final Cell cell = this.getPOICell(r, c);
		if (!this.isDate(cell))
			throw new IllegalArgumentException();
		return cell.getDateCellValue();
	}

	/** {@inheritDoc} */
	@Override
	public Double getDouble(final int r, final int c) {
		final Cell cell = this.getPOICell(r, c);
		if (!this.isDouble(cell))
			throw new IllegalArgumentException();

		return cell.getNumericCellValue();
	}

	/** {@inheritDoc} */
	@Override
	public String getFormula(final int r, final int c) {
		final Cell cell = this.getPOICell(r, c);
		if (cell.getCellType() != Cell.CELL_TYPE_FORMULA)
			throw new IllegalArgumentException();
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
		if (rowCount == 1 && this.getCellCount(0) == 0)
			rowCount = 0;
		return rowCount;
	}

	/** {@inheritDoc} */
	@Override
	public String getText(final int r, final int c) {
		final Cell cell = this.getPOICell(r, c);
		if (cell.getCellType() != Cell.CELL_TYPE_STRING)
			throw new IllegalArgumentException(
					"This cell does not contain text");
		return cell.getStringCellValue();
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
	protected Cell getPOICell(final int r, final int c) {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();
		if (r >= this.getRowCount() || c >= this.getCellCount(r))
			return null;
		
		if (r != this.curR || this.curRow == null) {
			this.curRow = this.sheet.getRow(r);
			assert this.curRow != null; // else this.curRow = this.sheet.createRow(r);
			this.curR = r;
		}
		Cell cell = this.curRow.getCell(c);
		assert cell != null; // else cell = this.curRow.createCell(c);
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