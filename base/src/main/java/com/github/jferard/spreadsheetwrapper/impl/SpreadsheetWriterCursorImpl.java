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
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * Implementation of the writer Cursor, using a writer.
 */
public class SpreadsheetWriterCursorImpl extends AbstractSpreadsheetWriterCell
implements SpreadsheetWriterCursor {

	/** the cursor */
	private final Cursor cursor;
	/** the writer */
	private final SpreadsheetWriter writer;

	/**
	 * Creates a cursor from the writer
	 *
	 * @param writer
	 *            the writer inside the cursor
	 */
	public SpreadsheetWriterCursorImpl(final SpreadsheetWriter writer) {
		super();
		this.writer = writer;
		this.cursor = new AbstractCursor(writer.getRowCount()) {
			/** {@inheritDoc} */
			@Override
			protected int getCellCount(final int r) {
				return writer.getCellCount(r);
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
		return this.writer.getCellContent(this.cursor.getR(),
				this.cursor.getC());
	}

	/** {@inheritDoc} */
	@Override
	public Date getDate() {
		return this.writer.getDate(this.cursor.getR(), this.cursor.getC());
	}

	/** {@inheritDoc} */
	@Override
	public Double getDouble() {
		return this.writer.getDouble(this.cursor.getR(), this.cursor.getC());
	}

	/** {@inheritDoc} */
	@Override
	public String getFormula() {
		return this.writer.getFormula(this.cursor.getR(), this.cursor.getC());
	}

	/** {@inheritDoc} */
	@Override
	public Integer getInteger() {
		return this.writer.getInteger(this.cursor.getR(), this.cursor.getC());
	}

	/** {@inheritDoc} */
	@Override
	public int getR() {
		return this.cursor.getR();
	}

	/** {@inheritDoc} */
	public int getRowCount() {
		return this.writer.getRowCount();
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/String getStyleName() {
		return this.writer.getStyleName(this.cursor.getR(), this.cursor.getC());
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleString() {
		return this.writer.getStyleString(this.cursor.getR(),
				this.cursor.getC());
	}

	/** {@inheritDoc} */
	@Override
	public String getText() {
		return this.writer.getText(this.cursor.getR(), this.cursor.getC());
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
	public Object setCellContent(final Object content) {
		return this.writer.setCellContent(this.cursor.getR(),
				this.cursor.getC(), content);
	}

	/** {@inheritDoc} */
	@Override
	public Object setCellContent(final Object content, final String styleName) {
		return this.writer.setCellContent(this.cursor.getR(),
				this.cursor.getC(), content, styleName);
	}

	/** {@inheritDoc} */
	@Override
	public Date setDate(final Date date) {
		return this.writer
				.setDate(this.cursor.getR(), this.cursor.getC(), date);
	}

	/** {@inheritDoc} */
	@Override
	public Double setDouble(final Number value) {
		return this.writer.setDouble(this.cursor.getR(), this.cursor.getC(),
				value);
	}

	/** {@inheritDoc} */
	@Override
	public String setFormula(final String text) {
		return this.writer.setFormula(this.cursor.getR(), this.cursor.getC(),
				text);
	}

	/** {@inheritDoc} */
	@Override
	public Integer setInteger(final Number value) {
		return this.writer.setInteger(this.cursor.getR(), this.cursor.getC(),
				value);
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyleName(final String styleName) {
		return this.writer.setStyleName(this.cursor.getR(), this.cursor.getC(),
				styleName);
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public boolean setStyleString(final String styleString) {
		return this.writer.setStyleString(this.cursor.getR(),
				this.cursor.getC(), styleString);
	}

	/** {@inheritDoc} */
	@Override
	public String setText(final String text) {
		return this.writer
				.setText(this.cursor.getR(), this.cursor.getC(), text);
	}

	/** {@inheritDoc} */
	@Override
	public Move up() {
		return this.cursor.up();
	}
}