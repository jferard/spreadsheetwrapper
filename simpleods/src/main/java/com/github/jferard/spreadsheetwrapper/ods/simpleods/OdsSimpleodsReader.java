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
import java.util.Date;
import java.util.Locale;
import java.util.TimeZone;

import org.simpleods.ObjectQueue;
import org.simpleods.Table;
import org.simpleods.TableCell;
import org.simpleods.TableRow;

import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetReader;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 */
class OdsSimpleodsReader extends AbstractSpreadsheetReader implements
		SpreadsheetReader {
	/** index of current row, -1 if none */
	private int curR;

	/** current row, null if none */
	private/*@Nullable*/TableRow curRow;

	/** *internal* table */
	private final Table table;

	/**
	 * @param table
	 *            *internal* table
	 */
	OdsSimpleodsReader(final Table table) {
		super();
		this.table = table;
		this.curR = -1;
		this.curRow = null;
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
			throw new IllegalArgumentException(String.format(
					"Unknown type of cell %s", type));
		}
		return result;
	}

	/** {@inheritDoc} */
	@Override
	public int getCellCount(final int r) {
		final int count;
		final TableRow row = (TableRow) this.table.getRows().get(r);
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
		if (simpleCell == null)
			return null;

		return simpleCell.getStyle();
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
}