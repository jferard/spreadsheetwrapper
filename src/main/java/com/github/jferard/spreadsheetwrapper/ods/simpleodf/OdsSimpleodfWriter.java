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
import java.util.List;
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

	private final Table table;

	/**
	 * @param table
	 *            the *internal* table
	 */
	OdsSimpleodfWriter(final Table table) {
		super(new OdsSimpleodfReader(table));
		this.table = table;
		this.preader = (OdsSimpleodfReader) this.reader;

	}

	/** {@inheritDoc} */
	@Override
	public String getStyleName(final int r, final int c) {
		final Cell cell = this.preader.getSimpleCell(r, c);
		return cell.getCellStyleName();
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleString(final int r, final int c) {
		final Cell cell = this.preader.getSimpleCell(r, c);
		return cell.getCellStyleName();
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
	public List<Object> removeCol(final int c) {
		final List<Object> retValue = this.getColContents(c);
		this.table.removeColumnsByIndex(c, 1);
		return retValue;
	}

	@Override
	public List<Object> removeRow(final int r) {
		final List<Object> retValue = this.getRowContents(r);
		this.table.removeRowsByIndex(r, 1);
		return retValue;
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public Boolean setBoolean(final int r, final int c, final Boolean value) {
		final Cell cell = this.preader.getSimpleCell(r, c);
		cell.setBooleanValue(value);
		return value;
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public Date setDate(final int r, final int c, final Date date) {
		final Cell cell = this.preader.getSimpleCell(r, c);
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
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public Double setDouble(final int r, final int c, final Number value) {
		final Cell cell = this.preader.getSimpleCell(r, c);
		final double retValue = value.doubleValue();
		cell.setDoubleValue(retValue);
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public String setFormula(final int r, final int c, final String formula) {
		final Cell cell = this.preader.getSimpleCell(r, c);
		cell.setFormula("=" + formula);
		return formula;
	}

	/** {@inheritDoc} */
	@Override
	public Integer setInteger(final int r, final int c, final Number value) {
		final Cell cell = this.preader.getSimpleCell(r, c);
		final int retValue = value.intValue();
		cell.setDoubleValue(Double.valueOf(retValue));
		cell.setFormatString(OdsSimpleodfWriter.INT_FORMAT_STR);
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c, final String styleName) {
		final Cell cell = this.preader.getSimpleCell(r, c);
		cell.setCellStyleName(styleName);
		return true;
	}

	/**
	 * {@inheritDoc}
	 *
	 * @return
	 */
	@Override
	public String setText(final int r, final int c, final String text) {
		final Cell cell = this.preader.getSimpleCell(r, c);
		cell.setStringValue(text);
		return text;
	}

}