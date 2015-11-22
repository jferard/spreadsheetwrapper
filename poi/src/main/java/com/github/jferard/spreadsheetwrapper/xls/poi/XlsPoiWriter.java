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
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 */
class XlsPoiWriter extends AbstractSpreadsheetWriter implements
		SpreadsheetWriter {
	/**
	 * COPY FROM JXL : The maximum number of columns
	 */
	private final static int MAX_COLUMNS = 256;

	/**
	 * COPY FROM JXL : The maximum number of rows excel allows in a worksheet
	 */
	private final static int MAX_ROWS_PER_SHEET = 65536;

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

	/**
	 * @param traitStyleHelper
	 */
	XlsPoiWriter(final Sheet sheet, final XlsPoiStyleHelper styleHelper,
			final/*@Nullable*/CellStyle dateCellStyle) {
		super(new XlsPoiReader(sheet, styleHelper));
		this.styleHelper = styleHelper;
		this.sheet = sheet;
		this.dateCellStyle = dateCellStyle;

	}

	/** {@inheritDoc} */
	@Override
	public void insertCol(final int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public void insertRow(final int r) {
		this.sheet.shiftRows(r, this.getRowCount(), 1);
		this.sheet.createRow(r);
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> removeCol(final int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> removeRow(final int r) {
		final List</*@Nullable*/Object> ret = this.getRowContents(r);
		final Row row = this.sheet.getRow(r);
		this.sheet.removeRow(row);
		this.sheet.shiftRows(r, this.getRowCount(), 1);
		return ret;
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
		final CellStyle cellStyle = this.styleHelper.getCellStyle(
				this.sheet.getWorkbook(), wrapperCellStyle);
		poiCell.setCellStyle(cellStyle);
		return true;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c, final String styleName) {
		final CellStyle cellStyle = this.styleHelper.getCellStyle(
				this.sheet.getWorkbook(), styleName);
		if (cellStyle == null)
			return false;
		else {
			final Cell cell = this.getOrCreatePOICell(r, c);
			cell.setCellStyle(cellStyle);
			return true;
		}
	}

	/** */
	@Override
	public String setText(final int r, final int c, final String text) {
		final Cell cell = this.getOrCreatePOICell(r, c);
		cell.setCellValue(text);
		return text;
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
		if (r >= XlsPoiWriter.MAX_ROWS_PER_SHEET
				|| c >= XlsPoiWriter.MAX_COLUMNS)
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

}