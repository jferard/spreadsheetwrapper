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

import java.util.Date;

import jxl.BooleanCell;
import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.FormulaCell;
import jxl.LabelCell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.biff.formula.FormulaException;
import jxl.write.Formula;

import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetReader;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 */
class XlsJxlReader extends AbstractSpreadsheetReader implements
SpreadsheetReader {
	private static Date getDate(final Cell cell) {
		if (cell instanceof DateCell)
			return ((DateCell) cell).getDate();

		throw new IllegalArgumentException();
	}

	/** index of the current row, -1 if none */
	private int curR;

	/** the current row, null if none */
	private Cell /*@Nullable*/[] curRow;

	/** *internal* sheet */
	private final Sheet sheet;

	/**
	 * @param sheet
	 *            *internal* sheet
	 */
	XlsJxlReader(final Sheet sheet) {
		super();
		this.sheet = sheet;
		this.curR = -1;
		this.curRow = null;
	}

	/** {@inheritDoc} */
	@Override
	public Boolean getBoolean(final int r, final int c) {
		final Cell cell = this.getJxlCell(r, c);
		if (cell instanceof BooleanCell)
			return ((BooleanCell) cell).getValue();

		throw new IllegalArgumentException();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Object getCellContent(final int r, final int c) {
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
		return this.sheet.getRow(r).length;
	}

	/** {@inheritDoc} */
	@Override
	public Date getDate(final int r, final int c) {
		final Cell cell = this.getJxlCell(r, c);
		return XlsJxlReader.getDate(cell);
	}

	/** {@inheritDoc} */
	@Override
	public Double getDouble(final int r, final int c) {
		final Cell cell = this.getJxlCell(r, c);
		if (cell instanceof NumberCell)
			return ((NumberCell) cell).getValue();

		throw new IllegalArgumentException();
	}

	/** {@inheritDoc} */
	@Override
	public String getFormula(final int r, final int c) {
		final Cell cell = this.getJxlCell(r, c);
		String formula;
		if (cell instanceof FormulaCell) // read
			try {
				formula = ((FormulaCell) cell).getFormula();
			} catch (final FormulaException e) {
				throw new IllegalArgumentException(e);
			}
		else if (cell instanceof Formula) // write
			formula = ((Formula) cell).getContents();
		else
			throw new IllegalArgumentException(cell.toString());

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
	public String getText(final int r, final int c) {
		final Cell cell = this.getJxlCell(r, c);
		if (!(cell instanceof LabelCell))
			throw new IllegalArgumentException(cell.toString());

		return ((LabelCell) cell).getString();
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
	protected Cell getJxlCell(final int r, final int c) {
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
}