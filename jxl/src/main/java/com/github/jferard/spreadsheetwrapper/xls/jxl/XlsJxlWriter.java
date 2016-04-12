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

import jxl.BooleanCell;
import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.FormulaCell;
import jxl.LabelCell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.biff.formula.FormulaException;
import jxl.format.CellFormat;
import jxl.write.DateTime;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.WritableCell;
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

	private static Date getDate(final Cell cell) {
		final Date ret;
		if (cell instanceof DateCell)
			ret = ((DateCell) cell).getDate();
		else
			ret = null;
		return ret;
	}

	/** current row index, -1 if none */
	private int curR;

	/** current row, null if none */
	private Cell /*@Nullable*/[] curRow;

	/** *internal* sheet */
	private final Sheet sheet;

	/** helper for style */
	private final XlsJxlStyleHelper styleHelper;

	/**
	 * @param sheet
	 *            *internal* sheet
	 */
	XlsJxlWriter(final Sheet sheet,
			final XlsJxlStyleHelper styleHelper) {
		this.sheet = sheet;
		this.styleHelper = styleHelper;
		this.curR = -1;
		this.curRow = null;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Boolean getBoolean(final int r, final int c) {
		final Cell cell = this.getJxlCell(r, c);
		final Boolean ret;
		if (cell instanceof BooleanCell)
			ret = ((BooleanCell) cell).getValue();
		else
			ret = null;
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Object getCellContent(final int r, final int c) { // NOPMD
		// by
		// Julien
		// on
		// 22/11/15
		// 06:30
		final/*@Nullable*/Cell cell = this.getJxlCell(r, c);
		if (cell == null)
			return null;

		if (cell instanceof FormulaCell) // read
			try {
				return ((FormulaCell) cell).getFormula();
			} catch (final FormulaException e) {
				throw new IllegalArgumentException(e);
			}
		else if (cell instanceof Formula) // write
			return ((Formula) cell).getContents();

		final CellType type = cell.getType();
		Object result;

		if (type == CellType.BOOLEAN)
			result = this.getBoolean(r, c);
		else if (type == CellType.DATE)
			result = this.getDate(r, c);
		else if (type == CellType.EMPTY)
			result = null;
		else if (type == CellType.ERROR)
			result = null;
		else if (type == CellType.LABEL)
			result = this.getText(r, c);
		else if (type == CellType.NUMBER) {
			final double value = ((NumberCell) cell).getValue();
			if (value == Math.rint(value))
				result = Integer.valueOf((int) value);
			else
				result = Double.valueOf(value);
		} else
			throw new IllegalArgumentException(String.format(
					"There is not support for this type of cell %s", type));
		return result;
	}

	/** {@inheritDoc} */
	@Override
	public int getCellCount(final int r) {
		if (r < 0)
			throw new IllegalArgumentException();
		
		final int ret;
		if (r < this.sheet.getRows())
			ret =  this.sheet.getRow(r).length;
		else
			ret = 0;
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Date getDate(final int r, final int c) {
		final Cell cell = this.getJxlCell(r, c);
		if (cell == null)
			return null;

		return XlsJxlWriter.getDate(cell);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Double getDouble(final int r, final int c) {
		final Cell cell = this.getJxlCell(r, c);
		final Double ret;
		if (cell instanceof NumberCell)
			ret = ((NumberCell) cell).getValue();
		else
			ret = null;
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getFormula(final int r, final int c) {
		final Cell cell = this.getJxlCell(r, c);
		if (cell == null)
			return null;

		String formula;
		if (cell instanceof FormulaCell) // read
			try {
				formula = ((FormulaCell) cell).getFormula(); // NOPMD by Julien
				// on 22/11/15
				// 06:28
			} catch (final FormulaException e) {
				throw new IllegalArgumentException(e);
			}
		else if (cell instanceof Formula) // write
			formula = ((Formula) cell).getContents();
		else
			formula = null;

		return formula;
	}

	/** {@inheritDoc} */
	@Override
	public String getName() {
		return this.sheet.getName();
	}

	/** {@inheritDoc} */
	@Override
	public int getRowCount() {
		return this.sheet.getRows();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/WrapperCellStyle getStyle(final int r, final int c) {
		final Cell cell = this.getJxlCell(r, c);
		return this.styleHelper.getStyle(cell);
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleName(final int r, final int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getText(final int r, final int c) {
		final Cell cell = this.getJxlCell(r, c);
		final String ret;
		if (cell instanceof LabelCell)
			ret = ((LabelCell) cell).getString();
		else
			ret = null;
		
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public boolean insertCol(final int c) {
		if (c < 0)
			throw new IllegalArgumentException();
		
		if (c < this.sheet.getColumns()) {
			((WritableSheet) this.sheet).insertColumn(c);
			return true;
		} else
			return false;
	}

	/** {@inheritDoc} */
	@Override
	public boolean insertRow(final int r) {
		if (r < 0)
			throw new IllegalArgumentException();
		
		if (r < this.sheet.getRows()) {
			((WritableSheet) this.sheet).insertRow(r);
			return true;
		} else
			return false;
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> removeCol(final int c) {
		if (c >= ((WritableSheet) this.sheet).getColumns())
			return null;
		
		final List</*@Nullable*/Object> ret = this.getColContents(c);
		((WritableSheet) this.sheet).removeColumn(c);
		return ret;
	}
	
	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> removeRow(final int r) {
		if (r >= this.getRowCount())
			return null;

		final List</*@Nullable*/Object> ret = this.getRowContents(r);
		((WritableSheet) this.sheet).removeRow(r);
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
		return this.styleHelper.setStyle(cell, wrapperCellStyle);
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c, final String styleName) {
		final WritableCell cell = this.getOrCreateJxlCell(r, c);
		return this.styleHelper.setStyleName(cell, styleName);
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
			((WritableSheet) this.sheet).addCell(cell);
			this.curR = -1; 
			this.curRow = null; 
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
	 * Simple optimization hidden inside a method.
	 *
	 * @param r
	 *            the row index
	 * @param c
	 *            the column index
	 * @return the cell
	 */
	private/*@Nullable*/Cell getJxlCell(final int r, final int c) {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();
		if (r >= this.getRowCount() || c >= this.getCellCount(r))
			return null;

		Cell cell;
		if (r != this.curR || this.curRow == null) {
			this.curRow = this.sheet.getRow(r);
			this.curR = r;
		}
		if (c < this.curRow.length)
			cell = this.curRow[c];
		else {
			cell = this.sheet.getCell(c, r);
			this.curRow = this.sheet.getRow(r);
		}
		return cell;
	}

	/**
	 * @param r
	 *            row index
	 * @param c
	 *            column index
	 * @return the *internal* cell
	 */
	private WritableCell getOrCreateJxlCell(final int r, final int c) {
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
			cell = ((WritableSheet) this.sheet).getWritableCell(c, r);
			this.curRow = this.sheet.getRow(r);
		}
		return cell;
	}
	
}