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
package com.github.jferard.spreadsheetwrapper.xls.jxl;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.logging.Logger;

import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.Accessor;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetWriterCursorImpl;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 */
public class XlsJxlDocumentWriter extends AbstractSpreadsheetDocumentWriter
implements SpreadsheetDocumentWriter {
	/** a Spreadsheet writer accessor by name and by index */
	private final Accessor<SpreadsheetWriter> accessor;
	/** *internal* workbook */
	private final WritableWorkbook writableWorkbook;

	/**
	 * @param workbook
	 *            *internal* workbook
	 */
	XlsJxlDocumentWriter(final Logger logger, final WritableWorkbook workbook) {
		super(logger, null);
		this.writableWorkbook = workbook;
		this.accessor = new Accessor<SpreadsheetWriter>();
		final WritableSheet[] sheets = this.writableWorkbook.getSheets();
		for (int n = 0; n < sheets.length; n++) {
			final WritableSheet sheet = sheets[n];
			final String name = sheet.getName();
			final SpreadsheetWriter reader = new XlsJxlWriter(sheet);
			this.accessor.put(name, n, reader);
		}
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter addSheet(final int index, final String sheetName) {
		if (index < 0 || index > this.writableWorkbook.getNumberOfSheets())
			throw new IndexOutOfBoundsException();
		if (this.writableWorkbook.getSheet(sheetName) != null)
			throw new IllegalArgumentException();
		final WritableSheet createSheet = this.writableWorkbook.createSheet(
				sheetName, index);
		return new XlsJxlWriter(createSheet);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter addSheet(final String sheetName) {
		return this.addSheet(this.writableWorkbook.getNumberOfSheets(),
				sheetName);
	}

	/**
	 * @throws SpreadsheetException
	 */
	@Override
	public void close() throws SpreadsheetException {
		try {
			this.writableWorkbook.close();
		} catch (final WriteException e) {
			throw new SpreadsheetException(e);
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriterCursor getNewCursorByIndex(final int index) {
		return new SpreadsheetWriterCursorImpl(this.getSpreadsheet(index));
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriterCursor getNewCursorByName(final String sheetName) {
		return new SpreadsheetWriterCursorImpl(this.getSpreadsheet(sheetName));
	}

	/** */
	@Override
	public int getSheetCount() {
		return this.writableWorkbook.getNumberOfSheets();
	}

	/** */
	@Override
	public List<String> getSheetNames() {
		return new ArrayList<String>(Arrays.asList(this.writableWorkbook
				.getSheetNames()));
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter getSpreadsheet(final int index) {
		final SpreadsheetWriter spreadsheet;
		if (this.accessor.hasByIndex(index))
			spreadsheet = this.accessor.getByIndex(index);
		else {
			final WritableSheet[] sheets = this.writableWorkbook.getSheets();
			if (index < 0 || index >= sheets.length)
				throw new NoSuchElementException(String.format(
						"No sheet at position %d", index));

			final WritableSheet sheet = sheets[index];
			spreadsheet = new XlsJxlWriter(sheet);
			this.accessor.put(sheet.getName(), index, spreadsheet);
		}
		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter getSpreadsheet(final String sheetName) {
		final SpreadsheetWriter spreadsheet;
		if (this.accessor.hasByName(sheetName))
			spreadsheet = this.accessor.getByName(sheetName);
		else
			spreadsheet = this.findSpreadsheet(sheetName);

		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	public void save() throws SpreadsheetException {
		try {
			this.writableWorkbook.write();
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public void saveAs(final OutputStream outputStream) {
		throw new UnsupportedOperationException();
	}

	private SpreadsheetWriter findSpreadsheet(final String sheetName) {
		final SpreadsheetWriter spreadsheet;

		final WritableSheet[] sheets = this.writableWorkbook.getSheets();
		for (int n = 0; n < sheets.length; n++) {
			final WritableSheet sheet = sheets[n];

			if (sheet.getName().equals(sheetName)) {
				spreadsheet = new XlsJxlWriter(sheet);
				this.accessor.put(sheetName, n, spreadsheet);
				return spreadsheet;
			}
		}
		throw new NoSuchElementException(String.format(
				"No %s sheet in writableWorkbook", sheetName));
	}
}
