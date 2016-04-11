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

import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public interface SpreadsheetReader {
			/**
			 * @param r
			 *            index of the row (0..)
			 * @param c
			 *            index of the column (0..)
			 * @return the boolean, null if r >= rowCount or c >= colCount or if
			 *         the cell doesn't exist or doesn't contain a boolean
			 * 
			 * @throws IllegalArgumentException
			 *             if r < 0 or c < 0
			 */
			/*@Nullable*/Boolean getBoolean(int r, int c);

			/**
			 * @param r
			 *            index of the row (0..)
			 * @param c
			 *            index of the column (0..)
			 * @return the content of the cell as Object, null if r >= rowCount
			 *         or c >= colCount or if the cell doesn't exist
			 * @throws IllegalArgumentException
			 *             if r < 0 or c < 0
			 */
			/*@Nullable*/Object getCellContent(int r, int c);

	/**
	 * @param r
	 *            index of the row (0..)
	 * @return the number of column of the row, 0 if the row doesn't exist. We can assert that {@code result >= 0}
	 * @throws IllegalArgumentException
	 *             if r < 0
	 */
	int getCellCount(int r);

	/**
	 * @param c
	 *            index of the column (0..)
	 * @return A {@code List} of {@code getRowCount()} elements (element is null
	 *         if a cell doesn't exist). If c >= 0, we can assert that
	 *         {@code s.getColContents(c).size() == s.getRowCount()}
	 * @throws IllegalArgumentException
	 *             if c < 0
	 */
	List</*@Nullable*/Object> getColContents(int c);

			/**
			 * @param r
			 *            index of the row (0..)
			 * @param c
			 *            index of the column (0..)
			 * @return the date, null if r >= rowCount or c >= colCount or if
			 *         the cell doesn't exist doesn't contain a date
			 * @throws IllegalArgumentException
			 *             if r < 0 or c < 0
			 */
			/*@Nullable*/Date getDate(int r, int c);

			/**
			 * @param r
			 *            index of the row (0..)
			 * @param c
			 *            index of the column (0..)
			 * @return the double, null if r >= rowCount or c >= colCount or if
			 *         the cell doesn't existdoesn't contain a double
			 * @throws IllegalArgumentException
			 *             if r < 0 or c < 0
			 */
			/*@Nullable*/Double getDouble(int r, int c);

			/**
			 * @param r
			 *            index of the row (0..)
			 * @param c
			 *            index of the column (0..)
			 * @return the formula as a {@code String}, null if r >= rowCount or
			 *         c >= colCount or if the cell doesn't exist or doesn't
			 *         contain a formula
			 * @throws IllegalArgumentException
			 *             if r < 0 or c < 0
			 */
			/*@Nullable*/String getFormula(int r, int c);

			/**
			 * @param r
			 *            index of the row (0..)
			 * @param c
			 *            index of the column (0..)
			 * @return the integer part of the number, null if r >= rowCount or
			 *         c >= colCount or if the cell doesn't exist or doesn't
			 *         contain an Integer
			 * @throws IllegalArgumentException
			 *             if r < 0 or c < 0
			 */
			/*@Nullable*/Integer getInteger(int r, int c);

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
	 * @return A {@code List} of {@code getColCount(r)} elements (element is
	 *         null if a cell doesn't exist). It means that the list may be
	 *         empty. If r >= 0, we can assert that
	 *         {@code s.getRowContents(r).size() == s.getCellCount(r)}
	 * @throws IllegalArgumentException
	 *             if r < 0
	 */
	List</*@Nullable*/Object> getRowContents(int r);

	/**
	 * @return the number of rows in the sheet. We can assert that {@code result >= 0}
	 */
	int getRowCount();

			/**
			 * @param r
			 *            index of the row (0..)
			 * @param c
			 *            index of the column (0..)
			 * @return the style of the cell : 1) null if r >=
			 *         rowCount or c >= colCount, or if the cell doesn't exist.
			 *         2) WrapperCellStyle.EMPTY if the cell has no style
			 *         3) a WrapperCellStyle object otherwise.
			 * @throws IllegalArgumentException
			 *             if r < 0 or c < 0
			 */
			/*@Nullable*/WrapperCellStyle getStyle(int r, int c);

			/**
			 * @param r
			 *            index of the row (0..)
			 * @param c
			 *            index of the column (0..)
			 * @return the style name of the cell : 1) null if r >=
			 *         rowCount or c >= colCount, or if the cell doesn't exist.
			 *         2) "" if the cell has no style
			 *         3) a String otherwise.
			 * @throws IllegalArgumentException
			 *             if r < 0 or c < 0
			 */
			/*@Nullable*/String getStyleName(int r, int c);

			/**
			 * @param r
			 *            index of the row (0..)
			 * @param c
			 *            index of the column (0..)
			 * @return the text, null if r >= rowCount or c >= colCount, or if the cell doesn't exist or dosen't contain a text. 
			 * @throws IllegalArgumentException
			 *             if r < 0 or c < 0
			 */
			/*@Nullable*/String getText(int r, int c);
}