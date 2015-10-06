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

import java.util.List;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public interface SpreadsheetDocumentReader {
	/**
	 * close resource
	 *
	 * @throws SpreadsheetException
	 */
	void close() throws SpreadsheetException;

	/**
	 * @param styleName
	 *            the name of the style
	 * @return the style
	 */
	/*@Nullable*/ WrapperCellStyle getCellStyle(String styleName);

	/**
	 * @param index
	 *            index of the sheet, table, ...
	 * @return the cursor.
	 * @throws IndexOutOfBoundsException
	 *             if the sheet, table, ... does not exist
	 */
	SpreadsheetReaderCursor getNewCursorByIndex(final int index);

	/**
	 * @param sheetName
	 *            name of the sheet, table, ...
	 * @return the cursor.
	 * @throws SpreadsheetException
	 *             if the sheet, table, ... does not exist
	 */
	SpreadsheetReaderCursor getNewCursorByName(final String sheetName)
			throws SpreadsheetException;

	/** @return Le nombre de feuilles */
	int getSheetCount();

	/** @return Les noms des feuilles */
	List<String> getSheetNames();

	/**
	 * @param index
	 *            index of the sheet, table, ...
	 * @return the reader.
	 */
	SpreadsheetReader getSpreadsheet(final int index);

	/**
	 * @param sheetName
	 *            name of the sheet, table, ...
	 * @return the reader.
	 */
	SpreadsheetReader getSpreadsheet(final String name);

	/**
	 * Gets a style string
	 *
	 * @param styleName
	 *            the name of the style
	 * @return the style string
	 * @deprecated
	 */
	@Deprecated
	/*@Nullable*/ String getStyleString(String styleName);
}