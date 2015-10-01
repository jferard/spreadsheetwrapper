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
import java.util.Map;
import java.util.Map.Entry;

import jxl.Cell;
import jxl.format.CellFormat;
import jxl.write.DateTime;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFeatures;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;

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

	private Map<String, WritableCellFormat> cellFormatByName;

	/**
	 * @param sheet
	 *            *internal* sheet
	 * @param cellFormatByName 
	 */
	XlsJxlWriter(final WritableSheet sheet, Map<String, WritableCellFormat> cellFormatByName) {
		super(new XlsJxlReader(sheet));
		this.sheet = sheet;
		this.cellFormatByName = cellFormatByName;
		this.curR = -1;
		this.curRow = null;
	}

	/**
	 * {@inheritDoc}
	 * 
	 * @return
	 */
	@Override
	public Boolean setBoolean(final int r, final int c, final Boolean value) {
		this.addCell(new jxl.write.Boolean(c, r, value));
		return value;
	}

	/**
	 * {@inheritDoc}
	 * 
	 * @return
	 */
	@Override
	public Date setDate(final int r, final int c, final Date date) {
		this.addCell(new DateTime(c, r, date));
		return date;
	}

	/** {@inheritDoc} */
	@Override
	public Double setDouble(final int r, final int c, final Number value) {
		double retValue = value.doubleValue();
		this.addCell(new jxl.write.Number(c, r, retValue));
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public String setFormula(final int r, final int c, final String fString) {
		if (fString == null || fString.equals(""))
			return "";

		Formula formula;
		final WritableCell xritableCell = this.getJxlCell(c, r);
		if (xritableCell == null)
			formula = new Formula(c, r, fString);
		else {
			final CellFormat cellFormat = xritableCell.getCellFormat();
			if (cellFormat == null)
				formula = new Formula(c, r, fString);
			else
				formula = new Formula(c, r, fString, cellFormat);
		}
		this.addCellWithStdErrWorkaround(formula);
		return formula.getContents();
	}

	/** {@inheritDoc} */
	@Override
	public Integer setInteger(final int r, final int c, final Number value) {
		int retValue = value.intValue();
		this.setDouble(c, r, new Double(retValue));
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public String setText(final int r, final int c, final String text) {
		this.addCell(new Label(c, r, text));
		return text;
	}

	private void addCell(final WritableCell cell) {
		try {
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
	private void addCellWithStdErrWorkaround(final WritableCell writableCell) {
		// WORKAROUND : jxl does not throw an error but a message on stderr
		System.err.flush(); // empties stderr
		final PrintStream originalErr = new PrintStream(System.err); // backup
		// of
		// stderr
		final OutputStream byteArrayOutputStream = new ByteArrayOutputStream();
		System.setErr(new PrintStream(byteArrayOutputStream)); // stderr >
		// bytearray

		this.addCell(writableCell);
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
	protected/*@Nullable*/WritableCell getJxlCell(final int r, final int c) {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();

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

	/** {@inheritDoc} */
	@Override
	public void insertCol(int c) {
		this.sheet.insertColumn(c);
	}

	/** {@inheritDoc} */
	@Override
	public void insertRow(int r) {
		this.sheet.insertRow(r);
	}

	/** {@inheritDoc} */
	@Override
	public List<Object> removeCol(int c) {
		List<Object> ret = this.getColContents(c);
		this.sheet.removeColumn(c);
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public List<Object> removeRow(int r) {
		List<Object> ret = this.getRowContents(r);
		this.sheet.removeRow(r);
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleName(int r, int c) {
		WritableCell jxlCell = this.getJxlCell(r, c);
		CellFormat cf = jxlCell.getCellFormat();
		for (Map.Entry<String, WritableCellFormat> entry : this.cellFormatByName.entrySet()) {
			if (entry.getValue().equals(cf))
				return entry.getKey();
		}
		return null;
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleString(int r, int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(int r, int c, String styleName) {
		if (this.cellFormatByName.containsKey(styleName)) {
			CellFormat format = this.cellFormatByName.get(styleName);
			WritableCell jxlCell = this.getJxlCell(r, c);
			jxlCell.setCellFormat(format);
			return true;
		} else
			return false;
	}

}