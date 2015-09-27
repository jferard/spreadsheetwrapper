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
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.odftoolkit.odfdom.doc.table.OdfTableCell;

import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReaderCursor;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public abstract class AbstractSpreadsheetReader implements SpreadsheetReader {

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
			Logger.getLogger(OdfTableCell.class.getName()).log(Level.SEVERE,
					message, e);
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
	public final Integer getInteger(final int r, final int c) {
		return this.getDouble(r, c).intValue();
	}

	/** {@inheritDoc} */
	@Override
	public final SpreadsheetReaderCursor getNewCursor() {
		return new SpreadsheetReaderCursorImpl(this);
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

	/** {@inheritDoc} */
	@Override
	public final String getStyleName(final int r, final int c) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public final String getStyleString(final int r, final int c) {
		throw new UnsupportedOperationException();
	}
}