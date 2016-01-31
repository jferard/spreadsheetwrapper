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
package com.github.jferard.spreadsheetwrapper.impl;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.github.jferard.spreadsheetwrapper.DataWrapper;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * An abstract writer that handles style methods.
 */
public abstract class AbstractSpreadsheetWriter implements SpreadsheetWriter {
	/**
	 */
	public AbstractSpreadsheetWriter() {
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriterCursor getNewCursor() {
		return new SpreadsheetWriterCursorImpl(this);
	}

	/** {@inheritDoc} */
	@Override
	public Boolean setBoolean(final int r, final int c, final Boolean value,
			final String styleName) {
		final Boolean ret = this.setBoolean(r, c, value);
		this.setStyleName(r, c, styleName);
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public Object setCellContent(final int r, final int c, final Object content) {
		if (content == null)
			throw new IllegalArgumentException();

		Object retValue;
		if (content instanceof String)
			retValue = this.setText(r, c, (String) content);
		else if (content instanceof Double)
			retValue = this.setDouble(r, c, (Double) content);
		else if (content instanceof Integer)
			retValue = this.setInteger(r, c, (Integer) content);
		else if (content instanceof Number)
			retValue = this.setDouble(r, c, (Number) content);
		else if (content instanceof Boolean)
			retValue = this.setBoolean(r, c, (Boolean) content);
		else if (content instanceof Date) {
			retValue = this.setDate(r, c, (Date) content);
		} else if (content instanceof Calendar) {
			retValue = this.setDate(r, c, ((Calendar) content).getTime());
		} else
			throw new IllegalArgumentException();
		return retValue;
	}

	/** {@inheritDoc} */
	@Override
	public Object setCellContent(final int r, final int c,
			final Object content, final String styleName) {
		final Object ret = this.setCellContent(r, c, content);
		this.setStyleName(r, c, styleName);
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public List<Object> setColContents(final int c, final List<Object> contents) {
		final List<Object> ret = new ArrayList<Object>(contents.size());
		final int rowCount = this.getRowCount();
		for (int r = 0; r < rowCount; r++) {
			ret.add(r, this.setCellContent(r, c, contents.get(r)));
		}
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public Date setDate(final int r, final int c, final Date date,
			final String styleName) {
		final Date ret = this.setDate(r, c, date);
		this.setStyleName(r, c, styleName);
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public Double setDouble(final int r, final int c, final Number value,
			final String styleName) {
		final Double ret = this.setDouble(r, c, value);
		this.setStyleName(r, c, styleName);
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public String setFormula(final int r, final int c, final String formula,
			final String styleName) {
		final String ret = this.setFormula(r, c, formula);
		this.setStyleName(r, c, styleName);
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public Integer setInteger(final int r, final int c, final Number value,
			final String styleName) {
		final Integer ret = this.setInteger(r, c, value);
		this.setStyleName(r, c, styleName);
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public final List<Object> setRowContents(final int r,
			final List<Object> contents) {
		final List<Object> ret = new ArrayList<Object>(contents.size());
		final int colCount = this.getCellCount(r);
		for (int c = 0; c < colCount; c++)
			ret.add(c, this.setCellContent(r, c, contents.get(c)));
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final int r, final int c, final String styleName) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public String setText(final int r, final int c, final String text,
			final String styleName) {
		final String ret = this.setText(r, c, text);
		this.setStyleName(r, c, styleName);
		return ret;
	}

	/** {@inheritDoc} */
	@Override
	public boolean writeDataFrom(final int r, final int c,
			final DataWrapper dataWrapper) {
		return dataWrapper.writeDataTo(this, r, c);
	}

	/**
	 * @param dateString
	 *            the date as a String (e.g "2015-01-01")
	 * @param format
	 *            the format
	 * @return the date, null if can't parse
	 */
	public static/*@Nullable*/Date parseString(final String dateString,
			final String format) {
		final SimpleDateFormat simpleFormat = new SimpleDateFormat(format,
				Locale.US);
		Date simpleDate;
		try {
			simpleDate = simpleFormat.parse(dateString);
		} catch (final ParseException e) {
			String message = e.getMessage();
			if (message == null)
				message = "???";
			Logger.getLogger(AbstractSpreadsheetWriter.class.getName()).log(
					Level.SEVERE, message, e);
			simpleDate = null;
		}
		return simpleDate;
	}

	/** {@inheritDoc} */
	@Override
	public final List</*@Nullable*/Object> getColContents(final int colIndex) {
		final int rowCount = this.getRowCount();
		final List</*@Nullable*/Object> cellContents = new ArrayList</*@Nullable*/Object>(
				rowCount);
		for (int r = 0; r < rowCount; r++) {
			cellContents.add(this.getCellContent(r, colIndex));
		}
		return cellContents;
	}

	/** {@inheritDoc} */
	@Override
	public final/*@Nullable*/Integer getInteger(final int r, final int c) {
		final Double value = this.getDouble(r, c);
		if (value == null)
			return null;
		return value.intValue();
	}

	/** {@inheritDoc} */
	@Override
	public final List</*@Nullable*/Object> getRowContents(final int rowIndex) {
		final int colCount = this.getCellCount(rowIndex);
		final List</*@Nullable*/Object> cellContents = new ArrayList</*@Nullable*/Object>(
				colCount);
		for (int c = 0; c < colCount; c++)
			cellContents.add(this.getCellContent(rowIndex, c));
		return cellContents;
	}	
}
