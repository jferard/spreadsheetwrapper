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

import com.github.jferard.spreadsheetwrapper.Cursor;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReaderCursor;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * Builds a cursor form the reader.
 */
public class SpreadsheetReaderCursorImpl implements SpreadsheetReaderCursor {

	/** the cursor */
	private final Cursor cursor;
	/** the reader */
	private final SpreadsheetReader reader;

	/**
	 * @param reader
	 *            the reader on the sheet
	 */
	public SpreadsheetReaderCursorImpl(final SpreadsheetReader reader) {
		this.reader = reader;
		this.cursor = new AbstractCursor(reader.getRowCount()) {
			/** {@inheritDoc} */
			@Override
			protected int getCellCount(final int r) {
				return reader.getCellCount(r);
			}
		};
	}

	/** {@inheritDoc} */
	@Override
	public Move crlf() {
		return this.cursor.crlf();
	}

	/** {@inheritDoc} */
	@Override
	public Move down() {
		return this.cursor.down();
	}

	/** {@inheritDoc} */
	@Override
	public Move first() {
		return this.cursor.first();
	}

	/** {@inheritDoc} */
	@Override
	public Move forceDown() {
		return this.cursor.forceDown();
	}

	/** {@inheritDoc} */
	@Override
	public Move forceRight() {
		return this.cursor.forceRight();
	}

	/** {@inheritDoc} */
	@Override
	public int getC() {
		return this.cursor.getC();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Object getCellContent() {
		return this.reader.getCellContent(this.getR(), this.getC());
	}

	/** {@inheritDoc} */
	@Override
	public Date getDate() {
		return this.reader.getDate(this.cursor.getR(), this.cursor.getC());
	}

	/** {@inheritDoc} */
	@Override
	public Double getDouble() {
		return this.reader.getDouble(this.cursor.getR(), this.cursor.getC());
	}

	/** {@inheritDoc} */
	@Override
	public String getFormula() {
		return this.reader.getFormula(this.cursor.getR(), this.cursor.getC());
	}

	/** {@inheritDoc} */
	@Override
	public Integer getInteger() {
		return this.reader.getInteger(this.cursor.getR(), this.cursor.getC());
	}

	/** {@inheritDoc} */
	@Override
	public int getR() {
		return this.cursor.getR();
	}

	/** {@inheritDoc} */
	public int getRowCount() {
		return this.reader.getRowCount();
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleName() {
		return this.reader.getStyle(this.cursor.getR(), this.cursor.getC());
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleString() {
		return this.reader.getStyleString(this.cursor.getR(),
				this.cursor.getC());
	}

	/** {@inheritDoc} */
	@Override
	public String getText() {
		return this.reader.getText(this.cursor.getR(), this.cursor.getC());
	}

	/** {@inheritDoc} */
	@Override
	public Move last() {
		return this.cursor.last();
	}

	/** {@inheritDoc} */
	@Override
	public Move left() {
		return this.cursor.left();
	}

	/** {@inheritDoc} */
	@Override
	public Move next() {
		return this.cursor.next();
	}

	/** {@inheritDoc} */
	@Override
	public Move previous() {
		return this.cursor.previous();
	}

	/** {@inheritDoc} */
	@Override
	public Move right() {
		return this.cursor.right();
	}

	/** {@inheritDoc} */
	@Override
	public Move set(final int r, final int c) {
		return this.cursor.set(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public Move up() {
		return this.cursor.up();
	}
}