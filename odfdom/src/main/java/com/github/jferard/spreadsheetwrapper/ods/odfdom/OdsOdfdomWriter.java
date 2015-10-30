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
package com.github.jferard.spreadsheetwrapper.ods.odfdom;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;

import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.doc.table.OdfTableCell;
import org.odftoolkit.odfdom.doc.table.OdfTableRow;
import org.odftoolkit.odfdom.dom.element.table.TableTableCellElementBase;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 */
class OdsOdfdomWriter extends AbstractSpreadsheetWriter implements
SpreadsheetWriter {
	/** index of current row, -1 if none */
	private int curR;

	/** current row, null if none */
	private/*@Nullable*/OdfTableRow curRow;

	private final OdfTable table;

	/**
	 * @param table
	 *            the *internal* sheet
	 */
	OdsOdfdomWriter(final OdfTable table) {
		super(new OdsOdfdomReader(table));
		this.table = table;
		this.curR = -1;
		this.curRow = null;
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleName(final int r, final int c) {
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
		return cell.getStyleName();
	}

	@Override
	public void insertCol(final int c) {
		this.table.insertColumnsBefore(c, 1);
	}

	@Override
	public void insertRow(final int r) {
		this.table.insertRowsBefore(r, 1);
	}

	@Override
	public List</*@Nullable*/ Object> removeCol(final int c) {
		final List</*@Nullable*/ Object> retValue = this.getColContents(c);
		this.table.removeColumnsByIndex(c, 1);
		return retValue;
	}

	@Override
	public List</*@Nullable*/ Object> removeRow(final int r) {
		final List</*@Nullable*/ Object> retValue = this.getRowContents(r);
		this.table.removeRowsByIndex(r, 1);
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public Boolean setBoolean(final int r, final int c, final Boolean value) {
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
		cell.setBooleanValue(value);
		return value;
	}

	/**
	 */
	@Override
	public Date setDate(final int r, final int c, final Date date) {
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
		final Calendar cal = Calendar.getInstance();
		cal.setTime(date);
		cell.setDateValue(cal); // for the hidden manipulations.
		final TableTableCellElementBase odfElement = cell.getOdfElement();
		final SimpleDateFormat simpleFormat = new SimpleDateFormat(
				"yyyy-MM-dd'T'HH:mm:ss", Locale.US);
		final String svalue = simpleFormat.format(date);
		odfElement.setOfficeDateValueAttribute(svalue);
		final Date retDate = new Date();
		retDate.setTime(date.getTime() / 1000 * 1000);
		return retDate;
	}

	/**
	 * @return
	 */
	@Override
	public Double setDouble(final int r, final int c, final Number value) {
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
		final double retValue = value.doubleValue();
		cell.setDoubleValue(retValue);
		return retValue;
	}

	/**  */
	@Override
	public String setFormula(final int r, final int c, final String formula) {
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
		cell.setFormula("=" + formula);
		cell.setStringValue("");
		return formula;
	}

	/**
	 * @return
	 */
	@Override
	public Integer setInteger(final int r, final int c, final Number value) {
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
		final int retValue = value.intValue();
		cell.setDoubleValue(Double.valueOf(retValue));
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c, final String styleName) {
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
		cell.getOdfElement().setStyleName(styleName);
		return true;
	}

	/**  */
	@Override
	public String setText(final int r, final int c, final String text) {
		final OdfTableCell cell = this.getOrCreateOdfCell(r, c);
		cell.setStringValue(text);
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
	protected OdfTableCell getOrCreateOdfCell(final int r, final int c) {
		if (r < 0 || c < 0)
			throw new IllegalArgumentException();
		if (r != this.curR || this.curRow == null) {
			this.curRow = this.table.getRowByIndex(r);
			this.curR = r;
		}
		return this.curRow.getCellByIndex(c);
	}
}