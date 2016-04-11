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
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

class SpreadsheetReaderImpl implements
		SpreadsheetReader {
	
	/** negative-decorator pattern : less methods than original */
	private SpreadsheetWriter spreadsheetWriter;

	SpreadsheetReaderImpl(SpreadsheetWriter spreadsheetWriter) {
		this.spreadsheetWriter = spreadsheetWriter;
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (!(obj instanceof SpreadsheetReaderImpl))
			return false;

		SpreadsheetReaderImpl other = (SpreadsheetReaderImpl) obj;
		return this.spreadsheetWriter.equals(other.spreadsheetWriter);
	}

	/** {@inheritDoc} */
	@Override
	public Boolean getBoolean(int r, int c) {
		return this.spreadsheetWriter.getBoolean(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public Object getCellContent(int r, int c) {
		return this.spreadsheetWriter.getCellContent(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public int getCellCount(int r) {
		return this.spreadsheetWriter.getCellCount(r);
	}

	/** {@inheritDoc} */
	@Override
	public List<Object> getColContents(int c) {
		return this.spreadsheetWriter.getColContents(c);
	}

	/** {@inheritDoc} */
	@Override
	public Date getDate(int r, int c) {
		return this.spreadsheetWriter.getDate(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public Double getDouble(int r, int c) {
		return this.spreadsheetWriter.getDouble(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public String getFormula(int r, int c) {
		return this.spreadsheetWriter.getFormula(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public Integer getInteger(int r, int c) {
		return this.spreadsheetWriter.getInteger(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public String getName() {
		return this.spreadsheetWriter.getName();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriterCursor getNewCursor() {
		return this.spreadsheetWriter.getNewCursor();
	}

	/** {@inheritDoc} */
	@Override
	public List<Object> getRowContents(int rowIndex) {
		return this.spreadsheetWriter.getRowContents(rowIndex);
	}

	/** {@inheritDoc} */
	@Override
	public int getRowCount() {
		return this.spreadsheetWriter.getRowCount();
	}

	/** {@inheritDoc} */
	@Override
	public WrapperCellStyle getStyle(int r, int c) {
		return this.spreadsheetWriter.getStyle(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleName(int r, int c) {
		return this.spreadsheetWriter.getStyleName(r, c);
	}

	/** {@inheritDoc} */
	@Override
	public String getText(int r, int c) {
		return this.spreadsheetWriter.getText(r, c);
	}

	@Override
	public int hashCode() {
		return this.spreadsheetWriter.hashCode();
	}
}