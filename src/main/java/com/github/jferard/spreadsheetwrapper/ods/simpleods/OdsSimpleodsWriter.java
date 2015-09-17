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
import java.util.TimeZone;

import org.simpleods.Table;
import org.simpleods.TableCell;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetWriter;

/**
 * writer for simpleods
 */
class OdsSimpleodsWriter extends AbstractSpreadsheetWriter implements
SpreadsheetWriter {

	/** the reader for delegation */
	private final OdsSimpleodsReader preader;

	/**
	 * @param table
	 *            *internal* sheet
	 */
	OdsSimpleodsWriter(final Table table) {
		super(new OdsSimpleodsReader(table));
		this.preader = (OdsSimpleodsReader) this.reader;

	}

	/** {@inheritDoc} */
	@Override
	public void setBoolean(final int r, final int c, final Boolean value) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public void setDate(final int r, final int c, final Date date) {
		final TableCell cell = this.preader.getSimpleCell(r, c);
		final TimeZone timeZone = TimeZone.getTimeZone("UTC");
		final Calendar cal = Calendar.getInstance(timeZone);
		cal.setTime(date);
		cell.setDateValue(cal);
	}

	/** {@inheritDoc} */
	@Override
	public void setDouble(final int r, final int c, final Double value) {
		final TableCell cell = this.preader.getSimpleCell(r, c);
		cell.setValue(value.toString());
		cell.setValueType(TableCell.STYLE_FLOAT);
	}

	/** {@inheritDoc} */
	@Override
	public void setFormula(final int r, final int c, final String formula) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public void setInteger(final int r, final int c, final Integer value) {
		final TableCell cell = this.preader.getSimpleCell(r, c);
		cell.setValue(value.toString());
		cell.setValueType(TableCell.STYLE_FLOAT);
	}

	/** {@inheritDoc} */
	@Override
	public void setText(final int r, final int c, final String text) {
		final TableCell cell = this.preader.getSimpleCell(r, c);
		cell.setText(text);
	}
}