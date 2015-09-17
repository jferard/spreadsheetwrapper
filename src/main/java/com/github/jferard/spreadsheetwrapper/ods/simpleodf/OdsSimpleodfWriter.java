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
package com.github.jferard.spreadsheetwrapper.ods.simpleodf;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;

import org.odftoolkit.odfdom.dom.element.table.TableTableCellElementBase;
import org.odftoolkit.simple.table.Cell;
import org.odftoolkit.simple.table.Table;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;

/**
 * A sheet writer for simple odf sheet.
 */
class OdsSimpleodfWriter extends AbstractSpreadsheetWriter implements
SpreadsheetWriter {
	/** format string for integers (internal : double) */
	private static final String INT_FORMAT_STR = "#";

	/** the reader for delegation */
	private final OdsSimpleodfReader preader;

	/**
	 * @param table
	 *            the *internal* table
	 */
	OdsSimpleodfWriter(final Table table) {
		super(new OdsSimpleodfReader(table));
		this.preader = (OdsSimpleodfReader) this.reader;

	}

	/** {@inheritDoc} */
	@Override
	public void setBoolean(final int r, final int c, final Boolean value) {
		final Cell cell = this.preader.getSimpleCell(r, c);
		cell.setBooleanValue(value);
	}

	/** {@inheritDoc} */
	@Override
	public void setDate(final int r, final int c, final Date date) {
		final Cell cell = this.preader.getSimpleCell(r, c);
		final Calendar cal = Calendar.getInstance();
		cal.setTime(date);
		cell.setDateValue(cal); // for the hidden manipulations.
		final TableTableCellElementBase odfElement = cell.getOdfElement();
		final SimpleDateFormat simpleFormat = new SimpleDateFormat(
				"yyyy-MM-dd'T'HH:mm:ss", Locale.US);
		final String svalue = simpleFormat.format(date);
		odfElement.setOfficeDateValueAttribute(svalue);
	}

	/** {@inheritDoc} */
	@Override
	public void setDouble(final int r, final int c, final Double value) {
		final Cell cell = this.preader.getSimpleCell(r, c);
		cell.setDoubleValue(value.doubleValue());
	}

	/** {@inheritDoc} */
	@Override
	public void setFormula(final int r, final int c, final String formula) {
		final Cell cell = this.preader.getSimpleCell(r, c);
		cell.setFormula("=" + formula);
	}

	/** {@inheritDoc} */
	@Override
	public void setInteger(final int r, final int c, final Integer value) {
		final Cell cell = this.preader.getSimpleCell(r, c);
		cell.setDoubleValue(value.doubleValue());
		cell.setFormatString(OdsSimpleodfWriter.INT_FORMAT_STR);
	}

	/** {@inheritDoc} */
	@Override
	public void setText(final int r, final int c, final String text) {
		final Cell cell = this.preader.getSimpleCell(r, c);
		cell.setStringValue(text);
	}
}