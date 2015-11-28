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

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.NoSuchElementException;

import jxl.Sheet;
import jxl.Workbook;
import jxl.write.WritableCellFormat;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReaderCursor;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.impl.Accessor;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetReaderCursorImpl;

/**
 */
public class XlsJxlDocumentReader implements SpreadsheetDocumentReader {
	/** accessor by name or index for the readers */
	private final Accessor<SpreadsheetReader> accessor;
	/** helper for style */
	private final XlsJxlStyleHelper styleHelper;

	/** *internal* workbook */
	private final Workbook workbook;

	/**
	 * @param workbook
	 *            *internal* workbook
	 */
	XlsJxlDocumentReader(final Workbook workbook,
			final XlsJxlStyleHelper styleHelper) {
		this.workbook = workbook;
		this.styleHelper = styleHelper;
		this.accessor = new Accessor<SpreadsheetReader>();
		final Sheet[] sheets = this.workbook.getSheets();
		for (int n = 0; n < sheets.length; n++) {
			final Sheet sheet = sheets[n];
			final String name = sheet.getName();
			final SpreadsheetReader reader = new XlsJxlReader(sheet, // NOPMD by Julien on 03/09/15 21:56
					styleHelper); 
			this.accessor.put(name, n, reader);
		}
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		this.workbook.close();
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
	public SpreadsheetReaderCursor getNewCursorByIndex(final int index) {
		return new SpreadsheetReaderCursorImpl(this.getSpreadsheet(index));
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetReaderCursor getNewCursorByName(final String sheetName) {
		return new SpreadsheetReaderCursorImpl(this.getSpreadsheet(sheetName));
	}

	/** {@inheritDoc} */
	@Override
	public int getSheetCount() {
		return this.workbook.getNumberOfSheets();
	}

	/** */
	@Override
	public List<String> getSheetNames() {
		return new ArrayList<String>(Arrays.asList(this.workbook
				.getSheetNames()));
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetReader getSpreadsheet(final int index) {
		final SpreadsheetReader spreadsheet;
		if (this.accessor.hasByIndex(index))
			spreadsheet = this.accessor.getByIndex(index);
		else {
			final Sheet[] sheets = this.workbook.getSheets();
			if (index < 0 || index >= sheets.length)
				throw new IndexOutOfBoundsException(String.format(
						"No sheet at position %d", index));

			final Sheet sheet = sheets[index];
			spreadsheet = new XlsJxlReader(sheet, this.styleHelper);
			this.accessor.put(sheet.getName(), index, spreadsheet);
		}
		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetReader getSpreadsheet(final String sheetName) {
		final SpreadsheetReader spreadsheet;
		if (this.accessor.hasByName(sheetName))
			spreadsheet = this.accessor.getByName(sheetName);
		else
			spreadsheet = this.findSpreadsheet(sheetName);

		return spreadsheet;
	}

	private SpreadsheetReader findSpreadsheet(final String sheetName) {
		final SpreadsheetReader spreadsheet;

		final Sheet[] sheets = this.workbook.getSheets();
		for (int n = 0; n < sheets.length; n++) {
			final Sheet sheet = sheets[n];

			if (sheet.getName().equals(sheetName)) {
				spreadsheet = new XlsJxlReader(sheet, this.styleHelper); // NOPMD
				// by
				// Julien
				// on
				// 03/09/15 21:57
				this.accessor.put(sheetName, n, spreadsheet);
				return spreadsheet;
			}
		}
		throw new NoSuchElementException(String.format(
				"No %s sheet in workbook", sheetName));
	}

}