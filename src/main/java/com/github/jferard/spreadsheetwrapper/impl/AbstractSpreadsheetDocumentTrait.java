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

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URL;
import java.net.URLConnection;
import java.net.UnknownServiceException;
import java.util.NoSuchElementException;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;

public abstract class AbstractSpreadsheetDocumentTrait<T> {

	/**
	 * @param outputURL
	 *            the URL to open for write
	 * @return the output stream on this URL
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public static OutputStream getOutputStream(final URL outputURL)
			throws IOException, FileNotFoundException {
		OutputStream outputStream;

		final URLConnection connection = outputURL.openConnection();
		connection.setDoOutput(true);
		try {
			outputStream = connection.getOutputStream(); // NOPMD by Julien on
			// 30/08/15 12:54
		} catch (final UnknownServiceException e) {
			outputStream = new FileOutputStream(outputURL.getPath());
		}
		return outputStream;
	}

	/**
	 * An accessor on readers/writers, by name and index.
	 */
	protected final Accessor<T> accessor;

	/**
	 * Creates an abstract spreadsheet document
	 */
	protected AbstractSpreadsheetDocumentTrait() {
		this.accessor = new Accessor<T>();
	}

	/**
	 * Adds a sheet to the document
	 *
	 * @param index
	 *            (0..n) the sheet is inserted before this index
	 * @param sheetName
	 *            the namee of the sheet
	 * @throws IndexOutOfBoundsException
	 *             if index < 0 or index > sheetCount
	 * @return the reader/writer
	 * @throws CantInsertElementInSpreadsheetException
	 */
	public T addSheet(final int index, final String sheetName)
			throws IndexOutOfBoundsException,
			CantInsertElementInSpreadsheetException {
		final int size = this.getSheetCount();
		if (index < 0 || index > size) // index == size is ok
			throw new IndexOutOfBoundsException();

		return this.addSheetWithCheckedIndex(index, sheetName);
	}

	/**
	 * Adds a sheet at the end of the workbook
	 *
	 * @param sheetName
	 *            the name of the sheet
	 * @return the reader/writer
	 * @throws CantInsertElementInSpreadsheetException
	 */
	public T addSheet(final String sheetName)
			throws CantInsertElementInSpreadsheetException {
		return this.addSheetWithCheckedIndex(this.getSheetCount(), sheetName);
	}

	/**
	 * @param sheetName
	 *            the name of the sheet to find
	 * @throws NoSuchElementException
	 *             if the workbook does not contains any sheet of this name
	 * @return the reader/writer
	 */
	public T getSpreadsheet(final String sheetName)
			throws NoSuchElementException {
		final T spreadsheet;
		if (this.accessor.hasByName(sheetName))
			spreadsheet = this.accessor.getByName(sheetName);
		else
			spreadsheet = this
					.findSpreadsheetAndCreateReaderOrWriter(sheetName);

		return spreadsheet;
	}

	/**
	 * @param index
	 *            the index of the new sheet (0..sheetCount)
	 * @param sheetName
	 *            the name of the sheet
	 * @return the reader/writer on thee sheet, table, ... inserted
	 * @throws CantInsertElementInSpreadsheetException
	 *             if cant inseert the spreadsheet in the workbook
	 */
	protected abstract T addSheetWithCheckedIndex(int index, String sheetName)
			throws CantInsertElementInSpreadsheetException;

	/**
	 * @param sheetName
	 *            the sheet, table, ... name
	 * @throws NoSuchElementException
	 *             if the workbook does not contains any sheet of this name
	 * @return the reader/writer
	 */
	protected abstract T findSpreadsheetAndCreateReaderOrWriter(String sheetName)
			throws NoSuchElementException;

	/**
	 * @return the number of sheets in the workbook
	 */
	protected abstract int getSheetCount();
}