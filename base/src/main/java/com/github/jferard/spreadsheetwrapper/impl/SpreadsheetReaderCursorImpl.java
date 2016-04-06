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
import com.github.jferard.spreadsheetwrapper.SpreadsheetReaderCursor;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * Builds a cursor form the reader.
 */
class SpreadsheetReaderCursorImpl implements SpreadsheetReaderCursor {

	/** the cursor */
	private final SpreadsheetWriterCursor cursor;


	/**
	 * @param reader
	 *            the reader on the sheet
	 */
	SpreadsheetReaderCursorImpl(final SpreadsheetWriterCursor cursor) {
		this.cursor = cursor;
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
		return this.cursor.getCellContent();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Date getDate() {
		return this.cursor.getDate();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Double getDouble() {
		return this.cursor.getDouble();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getFormula() {
		return this.cursor.getFormula();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/Integer getInteger() {
		return this.cursor.getInteger();
	}

	/** {@inheritDoc} */
	@Override
	public int getR() {
		return this.cursor.getR();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/WrapperCellStyle getStyle() {
		return this.cursor.getStyle();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getStyleName() {
		return this.cursor.getStyleName();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getText() {
		return this.cursor.getText();
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
	public Move up() { // NOPMD by Julien on 27/11/15 20:48
		return this.cursor.up();
	}
}