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

import java.util.Date;
import java.util.List;

import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;

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
	public final boolean createStyle(final String styleName,
			final String styleString) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public Boolean getBoolean(final int r, final int c) {
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
		return this.reader.getCellCount(r);
	}

	/** {@inheritDoc} */
	@Override
	public List</*@Nullable*/Object> getColContents(final int c) {
		return this.reader.getColContents(c);
	}

	/** {@inheritDoc} */
	@Override
	public Date getDate(final int r, final int c) {
		return this.reader.getDate(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public Double getDouble(final int r, final int c) {
		return this.reader.getDouble(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public String getFormula(final int r, final int c) {
		return this.reader.getFormula(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public Integer getInteger(final int r, final int c) {
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
	public String getStyle(final int r, final int c) {
		return this.reader.getStyle(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleString(final int r, final int c) {
		return this.reader.getStyleString(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public String getText(final int r, final int c) {
		return this.reader.getText(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public final void insertCol(final int r) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public final void insertRow(final int r) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public final void removeCol(final int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public final void removeRow(final int r) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public void setBoolean(final int r, final int c, final Boolean value,
			final String styleName) {
		this.setBoolean(r, c, value);
		this.setStyle(r, c, styleName);
	}

	/** {@inheritDoc} */
	@Override
	public final void setCellContent(final int r, final int c,
			final Object content) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public void setCellContent(final int r, final int c, final Object content,
			final String styleName) {
		this.setCellContent(r, c, content);
		this.setStyle(r, c, styleName);
	}

	/** {@inheritDoc} */
	@Override
	public final void setColContents(final int c, final List<Object> contents) {
		final int rowCount = this.getRowCount();
		for (int r = 0; r < rowCount; r++) {
			this.setCellContent(r, c, contents.get(r));
		}
	}

	/** {@inheritDoc} */
	@Override
	public void setDate(final int r, final int c, final Date date,
			final String styleName) {
		this.setDate(r, c, date);
		this.setStyle(r, c, styleName);
	}

	/** {@inheritDoc} */
	@Override
	public void setDouble(final int r, final int c, final Double value,
			final String styleName) {
		this.setDouble(r, c, value);
		this.setStyle(r, c, styleName);
	}

	/** {@inheritDoc} */
	@Override
	public void setFormula(final int r, final int c, final String formula,
			final String styleName) {
		this.setFormula(r, c, formula);
		this.setStyle(r, c, styleName);
	}

	/** {@inheritDoc} */
	@Override
	public void setInteger(final int r, final int c, final Integer value,
			final String styleName) {
		this.setInteger(r, c, value);
		this.setStyle(r, c, styleName);
	}

	/** {@inheritDoc} */
	@Override
	public final void setRowContents(final int r, final List<Object> contents) {
		final int colCount = this.getCellCount(r);
		for (int c = 0; c < colCount; c++)
			this.setCellContent(r, c, contents.get(c));
	}

	/** {@inheritDoc} */
	@Override
	public final boolean setStyle(final int r, final int c,
			final String styleName) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public final boolean setStyleString(final int r, final int c,
			final String styleString) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public void setText(final int r, final int c, final String text,
			final String styleName) {
		this.setText(r, c, text);
		this.setStyle(r, c, styleName);
	}
}
