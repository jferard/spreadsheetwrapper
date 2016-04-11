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

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.TimeZone;

import org.simpleods.ObjectQueue;
import org.simpleods.Table;
import org.simpleods.TableCell;
import org.simpleods.TableRow;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * writer for simpleods
 */
class OdsSimpleodsWriter extends AbstractSpreadsheetWriter
		implements SpreadsheetWriter {

	/** index of current row, -1 if none */
	private int curR;

	/** current row, null if none */
	private/*@Nullable*/TableRow curRow;

	/** the *internal* table */
	private final Table table;

	private OdsSimpleodsStyleHelper styleHelper;

	/**
	 * @param styleHelper
	 * @param table
	 *            *internal* sheet
	 */
	OdsSimpleodsWriter(OdsSimpleodsStyleHelper styleHelper, final Table table) {
		this.styleHelper = styleHelper;
		this.table = table;

	}

	/** {@inheritDoc} */
	@Override
	public Boolean getBoolean(final int r, final int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Object getCellContent(final int rowIndex,
			final int colIndex) {
		final TableCell cell = this.getSimpleCell(rowIndex, colIndex);
		if (cell == null)
			return null;

		Object result;
		final int type = cell.getValueType();
		switch (type) {
		case TableCell.STYLE_DATE:
			result = this.getDate(rowIndex, colIndex);
			break;
		case TableCell.STYLE_FLOAT:
			/*@SuppressWarnings("nullness")*/
			final double value = this.getDouble(rowIndex, colIndex);
			if (value == Math.rint(value))
				result = Integer.valueOf((int) value);
			else
				result = Double.valueOf(value);
			break;
		case TableCell.STYLE_STRING:
			result = cell.getText();
			break;
		default:
			throw new IllegalArgumentException(
					String.format("Unknown type of cell %s", type));
		}
		return result;
	}

	/** {@inheritDoc} */
	@Override
	public int getCellCount(final int r) {
		final ObjectQueue rows = this.table.getRows();
		if (r >= rows.size())
			return 0;

		final TableRow row = (TableRow) rows.get(r);
		final int count;
		if (row == null)
			count = 0;
		else {
			final ObjectQueue cells = row.getCells();
			count = cells == null ? 0 : cells.size();
		}
		return count;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Date getDate(final int r, final int c) {
		final TableCell cell = this.getSimpleCell(r, c);
		if (cell == null)
			return null;

		if (cell.getValueType() == TableCell.STYLE_DATE) {
			final String dateAsString = cell.getDateValue();
			final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd",
					Locale.US);
			final TimeZone zone = TimeZone.getDefault();
			try {
				final Date documentTrait = sdf.parse(dateAsString);
				return new Date(documentTrait.getTime() + zone.getRawOffset());
			} catch (final ParseException e) {
				throw new IllegalArgumentException(e);
			}
		}
		throw new IllegalArgumentException();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Double getDouble(final int r, final int c) {
		final TableCell cell = this.getSimpleCell(r, c);
		if (cell == null)
			return null;

		if (cell.getValueType() == TableCell.STYLE_FLOAT) {
			return Double.parseDouble(cell.getValue());
		}
		throw new IllegalArgumentException();
	}

	/** {@inheritDoc} */
	@Override
	public String getFormula(final int r, final int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public String getName() {
		return this.table.getName();
	}

	/** {@inheritDoc} */
	@Override
	public int getRowCount() {
		return this.table.getRows().size();
	}

	/** {@inheritDoc} */
	@Override
	public WrapperCellStyle getStyle(final int r, final int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getStyleName(final int r, final int c) {
		final TableCell simpleCell = this.getSimpleCell(r, c);
		return this.styleHelper.getStyleName(simpleCell);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getText(final int r, final int c) {
		final TableCell cell = this.getSimpleCell(r, c);
		if (cell == null)
			return null;

		if (cell.getValueType() == TableCell.STYLE_STRING) {
			return cell.getText();
		}
		throw new IllegalArgumentException();
	}

	/** {@inheritDoc} */
	@Override
	public boolean insertCol(final int c) {
		if (c < 0)
			throw new IllegalArgumentException();
		else
			throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public boolean insertRow(final int r) {
		if (r < 0)
			throw new IllegalArgumentException();
		else
			throw new UnsupportedOperationException();
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
		else
			throw new UnsupportedOperationException();
	}

	/**
	 * {@inheritDoc}
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
		cell.setValue(Double.toString(retValue));
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
		cell.setValue(Integer.toString(retValue));
		cell.setValueType(TableCell.STYLE_FLOAT);
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyle(final int r, final int c,
			final WrapperCellStyle wrapperStyle) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c,
			final String styleName) {
		final TableCell simpleCell = this.getOrCreateSimpleCell(r, c);
		return this.styleHelper.setStyle(simpleCell, styleName);
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
	private/*@Nullable*/TableCell getSimpleCell(final int r, final int c)
			throws IllegalArgumentException {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();
		if (r >= this.getRowCount() || c >= this.getCellCount(r))
			return null;

		if (r != this.curR || this.curRow == null) {
			final ObjectQueue rowsQueue = this.table.getRows();
			this.curRow = (TableRow) rowsQueue.get(r);
			assert this.curRow != null;
			// else final TableRow row = new TableRow();
			// if (rowsQueue.setAt(r, row))
			// this.curRow = row;
			// else {
			// this.curR = -1;
			// throw new IllegalArgumentException();
			// }
			this.curR = r;
		}
		return this.curRow.getCell(c);
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