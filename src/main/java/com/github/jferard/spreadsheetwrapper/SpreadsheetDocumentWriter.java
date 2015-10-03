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
package com.github.jferard.spreadsheetwrapper;

import java.io.File;
import java.io.OutputStream;
import java.net.URL;

/**
 * The SpreadsheetDocumentWriter class provides some basic methods for writing
 * in a spreadsheet file (xls or ods format)
 */
public interface SpreadsheetDocumentWriter extends SpreadsheetDocumentReader {
	/**
	 * Adds a sheet, table at the position index
	 *
	 * @param index
	 *            the position
	 * @param sheetName
	 *            the name of the new sheet, table, ...
	 * @return a writer on the sheet
	 * @throws CantInsertElementInSpreadsheetException
	 * @throws IndexOutOfBoundsException
	 */
	SpreadsheetWriter addSheet(int index, String sheetName)
			throws IndexOutOfBoundsException,
			CantInsertElementInSpreadsheetException;

	/**
	 * Adds a sheet, table at the end
	 *
	 * @param sheetName
	 *            the name of the new sheet, table, ...
	 * @return a writer on the sheet
	 * @throws CantInsertElementInSpreadsheetException
	 */
	SpreadsheetWriter addSheet(String sheetName)
			throws CantInsertElementInSpreadsheetException;

	/** {@inheritDoc} */
	@Override
	SpreadsheetWriterCursor getNewCursorByIndex(final int index);

	/** {@inheritDoc} */
	@Override
	SpreadsheetWriterCursor getNewCursorByName(final String sheetName)
			throws SpreadsheetException;

	/** {@inheritDoc} */
	@Override
	SpreadsheetWriter getSpreadsheet(final int index);

	/** {@inheritDoc} */
	@Override
	SpreadsheetWriter getSpreadsheet(final String name);

	/**
	 * Saves the current value
	 *
	 * @throws SpreadsheetException
	 *             if the value can't be saved
	 */
	void save() throws SpreadsheetException;

	/**
	 * Saves the current value
	 *
	 * @param outputFile
	 *            the destination file
	 * @throws SpreadsheetException
	 *             if the value can't be saved
	 */
	void saveAs(File outputFile) throws SpreadsheetException;

	/**
	 * Saves the current value
	 *
	 * @param outputStream
	 *            the destination stream
	 * @throws SpreadsheetException
	 *             if the value can't be saved
	 */
	void saveAs(OutputStream outputStream) throws SpreadsheetException;

	/**
	 * Saves the current value
	 *
	 * @param outputURL
	 *            the destination URL (local file is better)
	 * @throws SpreadsheetException
	 *             if the value can't be saved
	 * @deprecated use saveAs(outputStream)
	 */
	@Deprecated
	void saveAs(URL outputURL) throws SpreadsheetException;

	/**
	 * Creates a new style
	 *
	 * @param styleName
	 *            the name of the style
	 * @param styleString
	 *            the style string (format to be defined)
	 * @return false if fails
	 * @deprecated
	 */
	@Deprecated
	boolean createStyle(String styleName, String styleString);

	/**
	 * Creates a new style
	 *
	 * @param styleName
	 *            the name of the style
	 * @param cellStyle the style
	 * 
	 * @return false if fails
	 */
	boolean setStyle(String styleName, CellStyle cellStyle);
	
	/**
	 * Creates an existing style
	 *
	 * @param styleName
	 *            the name of the style
	 * @param styleString
	 *            the style string (format to be defined)
	 * @return false if fails
	 * @deprecated
	 */
	@Deprecated
	boolean updateStyle(String styleName, String styleString);
}