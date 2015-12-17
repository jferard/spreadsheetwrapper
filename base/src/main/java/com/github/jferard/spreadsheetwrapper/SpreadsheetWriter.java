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

/**
 *
 */
public interface SpreadsheetWriter extends SpreadsheetReader {

	/** {@inheritDoc} */
	@Override
	SpreadsheetWriterCursor getNewCursor();

	/**
	 * @param c
	 *            index (0..) of the column after the new column
	 */
	void insertCol(int c);

	/**
	 * @param r
	 *            index (0..) of the row after the new row
	 */
	void insertRow(int r);

	/**
	 * @param c
	 *            index (0..) of the column to be removed
	 */
	List</*@Nullable*/Object> removeCol(int c);

	/**
	 * @param r
	 *            index (0..) of the column to be removed
	 */
	List</*@Nullable*/Object> removeRow(int r);

	/**
	 * @param r
	 *            index of the row (0..)
	 * @param c
	 *            index of the column (0..)
	 * @param bool
	 *            the boolean value to put
	 */
	Boolean setBoolean(int r, int c, Boolean bool);

	/**
	 * @param r
	 *            index of the row (0..)
	 * @param c
	 *            index of the column (0..)
	 * @param bool
	 *            the boolean value to put
	 * @param styleName
	 *            the name of the style to use
	 */
	Boolean setBoolean(int r, int c, Boolean bool, String styleName);

	/**
	 * @param r
	 *            row index
	 * @param c
	 *            col index
	 * @param content
	 *            the content to put in the cell
	 */
	Object setCellContent(int r, int c, Object content);

	/**
	 * @param r
	 *            row index
	 * @param c
	 *            col index
	 * @param content
	 *            the content to put in the cell
	 * @param styleName
	 *            the style name
	 */
	Object setCellContent(int r, int c, Object content, String styleName);

	/**
	 * @param c
	 *            column index (0..)
	 * @param contents
	 *            the contents value to put in the cell
	 */
	List<Object> setColContents(int c, List<Object> contents);

	/**
	 * @param r
	 *            row index (0..)
	 * @param c
	 *            column index (0..)
	 * @param date
	 *            the date to put in the cell
	 */
	Date setDate(int r, int c, Date date);

	/**
	 * @param r
	 *            row index (0..)
	 * @param c
	 *            column index (0..)
	 * @param date
	 *            the date to put in the cell
	 * @param styleName
	 *            the style name
	 */
	Date setDate(int r, int c, Date date, String styleName);

	/**
	 * @param r
	 *            row index (0..)
	 * @param c
	 *            column index (0..)
	 * @param value
	 *            the double value to put in the cell
	 */
	Double setDouble(int r, int c, Number value);

	/**
	 * @param r
	 *            row index (0..)
	 * @param c
	 *            column index (0..)
	 * @param value
	 *            the double value to put in the cell
	 * @param styleName
	 *            the style name
	 */
	Double setDouble(int r, int c, Number value, String styleName);

	/**
	 * @param r
	 *            row index (0..)
	 * @param c
	 *            column index (0..)
	 * @param formula
	 *            the formulaa as a string. There is no test if such a formula
	 *            will have any meaning.
	 */
	String setFormula(int r, int c, String formula);

	/**
	 * @param r
	 *            row index (0..)
	 * @param c
	 *            column index (0..)
	 * @param formula
	 *            the formulaa as a string. There is no test if such a formula
	 *            will have any meaning.
	 * @param styleName
	 *            the style
	 */
	String setFormula(int r, int c, String formula, String styleName);

	/**
	 * @param r
	 *            row index (0..)
	 * @param c
	 *            column index (0..)
	 * @param value
	 *            the integer
	 */
	Integer setInteger(int r, int c, Number value);

	/**
	 * @param r
	 *            row index (0..)
	 * @param c
	 *            column index (0..)
	 * @param value
	 *            the integer
	 * @param styleName
	 *            the style
	 */
	Integer setInteger(int r, int c, Number value, String styleName);

	/**
	 * @param r
	 *            row index (0..)
	 * @param contents
	 *            the contents value to put in the cell
	 */
	List<Object> setRowContents(int r, List<Object> contents);

	/**
	 * @param r
	 *            row index (0..)
	 * @param c
	 *            column index (0..)
	 * @param style
	 *            the style (@see createStyle)
	 * @return false if failed
	 */
	boolean setStyle(int r, int c, WrapperCellStyle wrapperStyle);

	/**
	 * @param r
	 *            row index (0..)
	 * @param c
	 *            column index (0..)
	 * @param styleName
	 *            the name of the style (@see createStyle)
	 * @return false if failed
	 */
	boolean setStyleName(int r, int c, String styleName);

	/**
	 * @param r
	 *            row index (0..)
	 * @param c
	 *            column index (0..)
	 * @param text
	 *            the text to put in the cell
	 */
	String setText(int r, int c, String text);

	/**
	 * @param r
	 *            row index (0..)
	 * @param c
	 *            column index (0..)
	 * @param text
	 *            the text to put in the cell
	 * @param styleName
	 *            the style name
	 */
	String setText(int r, int c, String text, String styleName);

	/**
	 * @param r
	 *            row index (0..)
	 * @param c
	 *            column index (0..)
	 * @param dataWrapper
	 *            the data to put
	 */
	boolean writeDataFrom(int r, int c, DataWrapper dataWrapper);
}