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
	private final Table table;

	/**
	 * @param table
	 *            *internal* sheet
	 */
	OdsSimpleodsWriter(final Table table) {
		super(new OdsSimpleodsReader(table));
		this.table = table;
		this.preader = (OdsSimpleodsReader) this.reader;

	}

	/** {@inheritDoc} */
	@Override
	public String getStyleName(final int r, final int c) {
		final TableCell simpleCell = this.preader.getSimpleCell(r, c);
		return simpleCell.getStyle();
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleString(final int r, final int c) {
		throw new UnsupportedOperationException();
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
	public List<Object> removeCol(final int c) {
		throw new UnsupportedOperationException();
	}

	@Override
	public List<Object> removeRow(final int r) {
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
		final TableCell cell = this.preader.getSimpleCell(r, c);
		final TimeZone timeZone = TimeZone.getTimeZone("UTC");
		final Calendar cal = Calendar.getInstance(timeZone);
		cal.setTime(date);
		cell.setDateValue(cal);
		return date;
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public Double setDouble(final int r, final int c, final Number value) {
		final TableCell cell = this.preader.getSimpleCell(r, c);
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
		final TableCell cell = this.preader.getSimpleCell(r, c);
		final int retValue = value.intValue();
		cell.setValue(Integer.valueOf(retValue).toString());
		cell.setValueType(TableCell.STYLE_FLOAT);
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c, final String styleName) {
		final TableCell simpleCell = this.preader.getSimpleCell(r, c);
		simpleCell.setStyle(styleName);
		return true;
	}

	/** {@inheritDoc} */
	@Override
	public String setText(final int r, final int c, final String text) {
		final TableCell cell = this.preader.getSimpleCell(r, c);
		cell.setText(text);
		cell.setValue(text);
		return text;
	}

}