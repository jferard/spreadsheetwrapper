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

import java.util.Date;
import java.util.List;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public interface SpreadsheetReader {
	/**
	 * @param r
	 *            index of the row (0..)
	 * @param c
	 *            index of the column (0..)
	 * @return the boolean
	 * @throws IllegalArgumentException
	 *             if the cell doesn't exist or doesn't contain a boolean
	 */
	Boolean getBoolean(int r, int c);

	/**
	 * @param r
	 *            index of the row (0..)
	 * @param c
	 *            index of the column (0..)
	 * @return the content of the cell as Object, null il the cell doesn't exist
	 */
	/*@Nullable*/Object getCellContent(int r, int c);

	/**
	 * @param r
	 *            index of the row (0..)
	 * @return the number of column of the row
	 */
	int getCellCount(int r);

	/**
	 * @param c
	 *            index of the column (0..)
	 * @return the list of contents, null if a cell doesn't exist
	 */
	List</*@Nullable*/Object> getColContents(int c);

	/**
	 * @param r
	 *            index of the row (0..)
	 * @param c
	 *            index of the column (0..)
	 * @return the date
	 * @throws IllegalArgumentException
	 *             if the cell doesn't exist or doesn't contain a date
	 */
	Date getDate(int r, int c);

	/**
	 * @param r
	 *            index of the row (0..)
	 * @param c
	 *            index of the column (0..)
	 * @return the double
	 * @throws IllegalArgumentException
	 *             if the cell doesn't exist or doesn't contain a double
	 */
	Double getDouble(int r, int c);

	/**
	 * @param r
	 *            index of the row (0..)
	 * @param c
	 *            index of the column (0..)
	 * @return the formula
	 * @throws IllegalArgumentException
	 *             if the cell doesn't exist or doesn't contain a formula
	 */
	String getFormula(int r, int c);

	/**
	 * @param r
	 *            index of the row (0..)
	 * @param c
	 *            index of the column (0..)
	 * @return the integer part of the number
	 * @throws IllegalArgumentException
	 *             if the cell doesn't exist or doesn't contain a number
	 */
	Integer getInteger(int r, int c);

	/**
	 * @return the name of the reader (ie of the sheet, table, ...)
	 */
	String getName();

	/**
	 * @return a cursor on the sheet
	 */
	SpreadsheetReaderCursor getNewCursor();

	/**
	 * @param r
	 *            index of the row (0..)
	 * @return the list of contents, null if a cell doesn't exist
	 */
	List</*@Nullable*/Object> getRowContents(int rowIndex);

	/**
	 * @return the number of rows in the sheet, table, ...
	 */
	int getRowCount();

	/**
	 * @param r
	 *            index of the row (0..)
	 * @param c
	 *            index of the column (0..)
	 * @return the style name
	 * @throws IllegalArgumentException
	 *             if the cell doesn't exist
	 */
	/*@Nullable*/String getStyleName(int r, int c);

	/**
	 * @param r
	 *            index of the row (0..)
	 * @param c
	 *            index of the column (0..)
	 * @return the style string (@see createStyle)
	 * @throws IllegalArgumentException
	 *             if the cell doesn't exist
	 */
	String getStyleString(int r, int c);

	/**
	 * @param r
	 *            index of the row (0..)
	 * @param c
	 *            index of the column (0..)
	 * @return the text
	 * @throws IllegalArgumentException
	 *             if the cell doesn't exist or doesn't contain a text
	 */
	String getText(int r, int c);
}