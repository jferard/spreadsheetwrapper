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

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import com.github.jferard.spreadsheetwrapper.DataWrapper;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * An abstract writer that handles style methods.
 */
public abstract class AbstractSpreadsheetWriter implements SpreadsheetWriter {
	/** the reader for delegation */
	protected final SpreadsheetReader reader;

	/**
	 * @param reader
	 *            the reader
	 */
	public AbstractSpreadsheetWriter(final SpreadsheetReader reader) {
		this.reader = reader;
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Boolean getBoolean(final int r, final int c) {
		return this.reader.getBoolean(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Object getCellContent(final int r, final int c) {
		return this.reader.getCellContent(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public int getCellCount(final int r) {
		if (r < 0 || r >= this.getRowCount())
			throw new IllegalArgumentException();

		return this.reader.getCellCount(r);
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> getColContents(final int c) {
		return this.reader.getColContents(c);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Date getDate(final int r, final int c) {
		return this.reader.getDate(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Double getDouble(final int r, final int c) {
		return this.reader.getDouble(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getFormula(final int r, final int c) {
		return this.reader.getFormula(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Integer getInteger(final int r, final int c) {
		return this.reader.getInteger(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public String getName() {
		return this.reader.getName();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriterCursor getNewCursor() {
		return new SpreadsheetWriterCursorImpl(this);
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> getRowContents(final int r) {
		return this.reader.getRowContents(r);
	}

	/** {@inheritDoc} */
	@Override
	public int getRowCount() {
		return this.reader.getRowCount();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/WrapperCellStyle getStyle(final int r, final int c) {
		return this.reader.getStyle(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getStyleName(final int r, final int c) {
		return this.reader.getStyleName(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getText(final int r, final int c) {
		return this.reader.getText(r, c);
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

}
