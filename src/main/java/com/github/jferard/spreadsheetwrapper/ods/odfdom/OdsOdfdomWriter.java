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
import java.util.Locale;

import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.doc.table.OdfTableCell;
import org.odftoolkit.odfdom.dom.element.table.TableTableCellElementBase;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;

/**
 */
class OdsOdfdomWriter extends AbstractSpreadsheetWriter implements
		SpreadsheetWriter {
	/** reader for delegation */
	private final OdsOdfdomReader preader;

	/**
	 * @param table
	 *            the *internal* sheet
	 */
	OdsOdfdomWriter(final OdfTable table) {
		super(new OdsOdfdomReader(table));
		this.preader = (OdsOdfdomReader) this.reader;

	}

	/** {@inheritDoc} */
	@Override
	public void setBoolean(final int r, final int c, final Boolean value) {
		final OdfTableCell cell = this.preader.getOdfCell(r, c);
		cell.setBooleanValue(value);
	}

	/**
	 */
	@Override
	public void setDate(final int r, final int c, final Date date) {
		final OdfTableCell cell = this.preader.getOdfCell(r, c);
		final Calendar cal = Calendar.getInstance();
		cal.setTime(date);
		cell.setDateValue(cal); // for the hidden manipulations.
		final TableTableCellElementBase odfElement = cell.getOdfElement();
		final SimpleDateFormat simpleFormat = new SimpleDateFormat(
				"yyyy-MM-dd'T'HH:mm:ss", Locale.US);
		final String svalue = simpleFormat.format(date);
		odfElement.setOfficeDateValueAttribute(svalue);
	}

	/**
	 */
	@Override
	public void setDouble(final int r, final int c, final Double value) {
		final OdfTableCell cell = this.preader.getOdfCell(r, c);
		cell.setDoubleValue(value.doubleValue());
	}

	/**  */
	@Override
	public void setFormula(final int r, final int c, final String formula) {
		final OdfTableCell cell = this.preader.getOdfCell(r, c);
		cell.setFormula("=" + formula);
	}

	/**  */
	@Override
	public void setInteger(final int r, final int c, final Integer value) {
		final OdfTableCell cell = this.preader.getOdfCell(r, c);
		cell.setDoubleValue(value.doubleValue());
	}

	/**  */
	@Override
	public void setText(final int r, final int c, final String text) {
		final OdfTableCell cell = this.preader.getOdfCell(r, c);
		cell.setStringValue(text);
	}
}