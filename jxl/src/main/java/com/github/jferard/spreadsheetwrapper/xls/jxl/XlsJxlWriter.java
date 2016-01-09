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
package com.github.jferard.spreadsheetwrapper.xls.jxl;

import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.Date;
import java.util.List;

import jxl.Cell;
import jxl.format.CellFormat;
import jxl.write.DateTime;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.xls.XlsConstants;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 */
class XlsJxlWriter extends AbstractSpreadsheetWriter implements
SpreadsheetWriter {

	/** current row index, -1 if none */
	private int curR;

	/** current row, null if none */
	private Cell /*@Nullable*/[] curRow;

	/** *internal* sheet */
	private final WritableSheet sheet;

	/** helper for style */
	private final XlsJxlStyleHelper styleHelper;

	/**
	 * @param sheet
	 *            *internal* sheet
	 */
	public XlsJxlWriter(final WritableSheet sheet,
			final XlsJxlStyleHelper styleHelper) {
		super(new XlsJxlReader(sheet, styleHelper));
		this.sheet = sheet;
		this.styleHelper = styleHelper;
		this.curR = -1;
		this.curRow = null;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getStyleName(final int r, final int c) {
		return this.reader.getStyleName(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public void insertCol(final int c) {
		this.sheet.insertColumn(c);
	}

	/** {@inheritDoc} */
	@Override
	public void insertRow(final int r) {
		this.sheet.insertRow(r);
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> removeCol(final int c) {
		final List</*@Nullable*/Object> ret = this.getColContents(c);
		this.sheet.removeColumn(c);
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> removeRow(final int r) {
		final List</*@Nullable*/Object> ret = this.getRowContents(r);
		this.sheet.removeRow(r);
		return ret;
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public Boolean setBoolean(final int r, final int c, final Boolean value) {
		this.addCell(r, c, new jxl.write.Boolean(c, r, value));
		return value;
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public Date setDate(final int r, final int c, final Date date) {
		this.addCell(r, c, new DateTime(c, r, date));
		return date;
	}

	/** {@inheritDoc} */
	@Override
	public Double setDouble(final int r, final int c, final Number value) {
		final double retValue = value.doubleValue();
		this.addCell(r, c, new jxl.write.Number(c, r, retValue));
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public String setFormula(final int r, final int c, final String fString) {
		if (fString == null || fString.equals(""))
			return "";

		Formula formula;
		// final WritableCell writableCell = this.getJxlCell(c, r);
		// if (writableCell == null)
		formula = new Formula(c, r, fString);
		// else {
		// final CellFormat cellFormat = writableCell.getCellFormat();
		// if (cellFormat == null)
		// formula = new Formula(c, r, fString);
		// else
		// formula = new Formula(c, r, fString, cellFormat);
		// }
		this.addCellWithStdErrWorkaround(r, c, formula);
		return formula.getContents();
	}

	/** {@inheritDoc} */
	@Override
	public Integer setInteger(final int r, final int c, final Number value) {
		final int retValue = value.intValue();
		this.setDouble(r, c, new Double(retValue));
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyle(final int r, final int c,
			final WrapperCellStyle wrapperCellStyle) {
		final WritableCell cell = this.getOrCreateJxlCell(r, c);
		if (cell == null)
			return false;

		this.styleHelper.setWrapperCellStyle(cell, wrapperCellStyle);
		return true;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c, final String styleName) {
		final WritableCellFormat cellFormat = this.styleHelper
				.getCellFormat(styleName);
		if (cellFormat == null)
			return false;
		else {
			final WritableCell cell = this.getOrCreateJxlCell(r, c);
			cell.setCellFormat(cellFormat);
			return true;
		}
	}

	/** {@inheritDoc} */
	@Override
	public String setText(final int r, final int c, final String text) {
		this.addCell(r, c, new Label(c, r, text));
		return text;
	}

	private void addCell(final int r, final int c, final WritableCell cell) {
		try {
			final WritableCell oldCell = this.getOrCreateJxlCell(r, c);
			final CellFormat oldFormat = oldCell.getCellFormat();
			if (oldFormat != null)
				cell.setCellFormat(oldFormat);
			this.sheet.addCell(cell);
		} catch (final RowsExceededException e) {
			throw new IllegalArgumentException(e);
		} catch (final WriteException e) {
			throw new IllegalArgumentException(e);
		} catch (final ArrayIndexOutOfBoundsException e) {
			throw new IllegalArgumentException(e);
		}
	}

	/** Adds a cell */
	private void addCellWithStdErrWorkaround(final int r, final int c,
			final WritableCell writableCell) {
		// WORKAROUND : jxl does not throw an error but a message on stderr
		System.err.flush(); // empties stderr
		final PrintStream originalErr = new PrintStream(System.err); // backup
		// of
		// stderr
		final OutputStream byteArrayOutputStream = new ByteArrayOutputStream();
		System.setErr(new PrintStream(byteArrayOutputStream)); // stderr >
		// bytearray

		this.addCell(r, c, writableCell);
		System.err.flush(); // empties stderr
		System.setErr(originalErr); // original stderr
		final String errMsg = byteArrayOutputStream.toString();
		if (!errMsg.isEmpty())
			throw new IllegalArgumentException(errMsg);
	}

	/**
	 * @param r
	 *            row index
	 * @param c
	 *            column index
	 * @return the *internal* cell
	 */
	protected WritableCell getOrCreateJxlCell(final int r, final int c) {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();
		if (r >= XlsConstants.MAX_ROWS_PER_SHEET
				|| c >= XlsConstants.MAX_COLUMNS)
			throw new UnsupportedOperationException();

		WritableCell cell;
		if (r != this.curR || this.curRow == null) {
			this.curRow = this.sheet.getRow(r);
			this.curR = r;
		}
		if (c < this.curRow.length)
			cell = (WritableCell) this.curRow[c];
		else {
			cell = this.sheet.getWritableCell(c, r);
			this.curRow = this.sheet.getRow(r);
		}
		return cell;
	}
}