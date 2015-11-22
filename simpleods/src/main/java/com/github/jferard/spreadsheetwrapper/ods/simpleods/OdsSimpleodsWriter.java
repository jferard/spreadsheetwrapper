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
package com.github.jferard.spreadsheetwrapper.ods.simpleods;

import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.TimeZone;

import org.simpleods.ObjectQueue;
import org.simpleods.Table;
import org.simpleods.TableCell;
import org.simpleods.TableRow;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * writer for simpleods
 */
class OdsSimpleodsWriter extends AbstractSpreadsheetWriter implements
SpreadsheetWriter {

	/** index of current row, -1 if none */
	private int curR;

	/** current row, null if none */
	private/*@Nullable*/TableRow curRow;

	/** the *internal* table */
	private final Table table;

	/**
	 * @param table
	 *            *internal* sheet
	 */
	OdsSimpleodsWriter(final Table table) {
		super(new OdsSimpleodsReader(table));
		this.table = table;

	}

	/** {@inheritDoc} */
	@Override
	public String getStyleName(final int r, final int c) {
		final TableCell simpleCell = this.getOrCreateSimpleCell(r, c);
		return simpleCell.getStyle();
	}

	/** {@inheritDoc} */
	@Override
	public void insertCol(final int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public void insertRow(final int r) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> removeCol(final int c) {
		throw new UnsupportedOperationException();
	}

	@Override
	public List</*@Nullable*/Object> removeRow(final int r) {
		throw new UnsupportedOperationException();
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public Boolean setBoolean(final int r, final int c, final Boolean value) {
		throw new UnsupportedOperationException();
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public Date setDate(final int r, final int c, final Date date) {
		final TableCell cell = this.getOrCreateSimpleCell(r, c);
		final TimeZone timeZone = TimeZone.getTimeZone("UTC");
		final Calendar cal = Calendar.getInstance(timeZone);
		cal.setTime(date);
		cell.setDateValue(cal);
		final Date retDate = new Date();
		retDate.setTime(date.getTime() / 86400000 * 86400000);
		return retDate;
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public Double setDouble(final int r, final int c, final Number value) {
		final TableCell cell = this.getOrCreateSimpleCell(r, c);
		final double retValue = value.doubleValue();
		cell.setValue(Double.valueOf(retValue).toString());
		cell.setValueType(TableCell.STYLE_FLOAT);
		return retValue;
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public String setFormula(final int r, final int c, final String formula) {
		throw new UnsupportedOperationException();
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public Integer setInteger(final int r, final int c, final Number value) {
		final TableCell cell = this.getOrCreateSimpleCell(r, c);
		final int retValue = value.intValue();
		cell.setValue(Integer.valueOf(retValue).toString());
		cell.setValueType(TableCell.STYLE_FLOAT);
		return retValue;
	}

	@Override
	public boolean setStyle(final int r, final int c,
			final WrapperCellStyle wrapperStyle) {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();

		return true;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c, final String styleName) {
		final TableCell simpleCell = this.getOrCreateSimpleCell(r, c);
		simpleCell.setStyle(styleName);
		return true;
	}

	/** {@inheritDoc} */
	@Override
	public String setText(final int r, final int c, final String text) {
		final TableCell cell = this.getOrCreateSimpleCell(r, c);
		cell.setText(text);
		cell.setValue(text);
		return text;
	}

	/**
	 * @param r
	 *            row index
	 * @param c
	 *            column index
	 * @return the *internal* cell
	 * @throws IllegalArgumentException
	 *             if the row does not exist
	 */
	protected TableCell getOrCreateSimpleCell(final int r, final int c)
			throws IllegalArgumentException {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();

		if (r != this.curR || this.curRow == null) {
			final ObjectQueue rowsQueue = this.table.getRows();
			this.curRow = (TableRow) rowsQueue.get(r);
			if (this.curRow == null) {
				final TableRow row = new TableRow();
				if (rowsQueue.setAt(r, row))
					this.curRow = row;
				else {
					this.curR = -1;
					throw new IllegalArgumentException();
				}
			}
			this.curR = r;
		}
		return this.curRow.getCell(c);
	}

}