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

import jxl.Sheet;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WriteException;

import com.github.jferard.spreadsheetwrapper.Accessor;
import com.github.jferard.spreadsheetwrapper.Output;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetWriterCursorImpl;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/**
 */
class XlsJxlDocumentWriter extends AbstractSpreadsheetDocumentWriter
implements SpreadsheetDocumentWriter {
	/** a Spreadsheet writer accessor by name and by index */
	private final Accessor<SpreadsheetWriter> accessor;

	/** helper for style */
	private final XlsJxlStyleHelper styleHelper;

	/** union for *internal* workbook */
	private final JxlWorkbook jxlWorkbook;

	/**
	 * @param workbook
	 *            *internal* workbook
	 */
	XlsJxlDocumentWriter(final Logger logger,
			final XlsJxlStyleHelper styleHelper, final JxlWorkbook writableWorkbook) {
		super(logger, new Output());
		this.styleHelper = styleHelper;
		this.jxlWorkbook = writableWorkbook;
		this.accessor = new Accessor<SpreadsheetWriter>();
		final Sheet[] sheets = this.jxlWorkbook.getSheets();
		for (int n = 0; n < sheets.length; n++) {
			final Sheet sheet = sheets[n];
			final String name = sheet.getName();
			final SpreadsheetWriter reader = new XlsJxlWriter(sheet, // NOPMD by
					// Julien
					// on
					// 28/11/15
					// 16:09
					styleHelper);
			this.accessor.put(name, n, reader);
  		}
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter addSheet(final int index, final String sheetName) {
		if (index < 0 || index > this.jxlWorkbook.getNumberOfSheets())
			throw new IndexOutOfBoundsException();
		return this.addSheetWithCheckedIndex(index, sheetName);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter addSheet(final String sheetName) {
		return this.addSheet(this.jxlWorkbook.getNumberOfSheets(),
				sheetName);
	}

	/**
	 * @throws SpreadsheetException
	 */
	@Override
	public void close() throws SpreadsheetException {
		try {
			this.jxlWorkbook.close();
		} catch (final WriteException e) {
			throw new SpreadsheetException(e);
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public WrapperCellStyle getCellStyle(final String styleName) {
		final WritableCellFormat cellFormat = this.styleHelper
				.getCellFormat(styleName);
		if (cellFormat == null)
			return WrapperCellStyle.EMPTY;

		return this.styleHelper.toWrapperCellStyle(cellFormat);
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
		return this.jxlWorkbook.getNumberOfSheets();
	}

	/** */
	@Override
	public List<String> getSheetNames() {
		return new ArrayList<String>(Arrays.asList(this.jxlWorkbook
				.getSheetNames()));
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter getSpreadsheet(final int index) {
		final SpreadsheetWriter spreadsheet;
		if (this.accessor.hasByIndex(index))
			spreadsheet = this.accessor.getByIndex(index);
		else {
			final Sheet[] sheets = this.jxlWorkbook.getSheets();
			if (index < 0 || index >= sheets.length)
				throw new IndexOutOfBoundsException(String.format(
						"No sheet at position %d", index));

			final Sheet sheet = sheets[index];
			spreadsheet = new XlsJxlWriter(sheet, this.styleHelper);
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
			spreadsheet = this.findSpreadsheetNotYetInAccessor(sheetName);

		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	public void save() throws SpreadsheetException {
		try {
			this.jxlWorkbook.write();
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public void saveAs(final OutputStream outputStream) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyle(final String styleName,
			final WrapperCellStyle wrapperCellStyle) {
		final WritableCellFormat cellFormat = this.styleHelper
				.toCellFormat(wrapperCellStyle);
		this.styleHelper.putCellStyle(styleName, cellFormat);
		return true;
	}

	private SpreadsheetWriter findSpreadsheetNotYetInAccessor(
			final String sheetName) {
		final SpreadsheetWriter spreadsheet;

		final Sheet[] sheets = this.jxlWorkbook.getSheets();
		for (int n = 0; n < sheets.length; n++) {
			final Sheet sheet = sheets[n];

			if (sheet.getName().equals(sheetName)) {
				spreadsheet = new XlsJxlWriter(sheet, this.styleHelper); // NOPMD
				// by
				// Julien
				// on
				// 28/11/15
				// 16:09
				this.accessor.put(sheetName, n, spreadsheet);
				return spreadsheet;
			}
		}
		throw new NoSuchElementException(String.format(
				"No %s sheet in jxlWorkbook", sheetName));
	}

	@Override
	protected SpreadsheetWriter addSheetWithCheckedIndex(int index,
			String sheetName) {
		if (this.jxlWorkbook.getSheet(sheetName) != null)
			throw new IllegalArgumentException();
		final WritableSheet createSheet = this.jxlWorkbook.createSheet(
				sheetName, index);
		return new XlsJxlWriter(createSheet, this.styleHelper);
	}

	@Override
	protected SpreadsheetWriter findSpreadsheetAndCreateReaderOrWriter(
			String sheetName) throws NoSuchElementException {
		// TODO Auto-generated method stub
		return null;
	}
	
//	private SpreadsheetReader findSpreadsheet(final String sheetName) {
//		final SpreadsheetReader spreadsheet;
//
//		final Sheet[] sheets = this.workbook.getSheets();
//		for (int n = 0; n < sheets.length; n++) {
//			final Sheet sheet = sheets[n];
//
//			if (sheet.getName().equals(sheetName)) {
//				spreadsheet = new XlsJxlReader(sheet, this.styleHelper); // NOPMD
//				// by
//				// Julien
//				// on
//				// 03/09/15 21:57
//				this.accessor.put(sheetName, n, spreadsheet);
//				return spreadsheet;
//			}
//		}
//		throw new NoSuchElementException(String.format(
//				"No %s sheet in workbook", sheetName));
//	}
}
