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

public interface SpreadsheetWriterCursor extends SpreadsheetReaderCursor {
	/**
	 * Set a value at the current position
	 *
	 * @param content
	 *            the value
	 */
	Object setCellContent(Object content);

	/**
	 * Set a value at the current position
	 *
	 * @param content
	 *            the value
	 * @param styleName
	 *            the name of the format
	 */
	Object setCellContent(Object content, String styleName);

	/**
	 * Set a date at the current position
	 *
	 * @param date
	 *            the date
	 */
	Date setDate(Date date);

	/**
	 * Set a date at the current position
	 *
	 * @param date
	 *            the date
	 * @param styleName
	 *            the name of the format
	 */
	Date setDate(Date date, String styleName);

	/**
	 * set a Double at the current position
	 *
	 * @param value
	 *            Double
	 */
	Double setDouble(Number value);

	/**
	 * set a Double at the current position
	 *
	 * @param value
	 *            Double
	 * @param styleName
	 *            the name of the format
	 */
	Double setDouble(Number value, String styleName);

	/**
	 * Set a formula at the current position
	 *
	 * @param formula
	 *            the formula text (english)
	 */
	String setFormula(String formula);

	/**
	 * Set a formula at the current position
	 *
	 * @param formula
	 *            the formula text (english)
	 * @param styleName
	 *            the name of the format
	 */
	String setFormula(String formula, String styleName);

	/**
	 * Set an integer at the current position
	 *
	 * @param value
	 *            the integer
	 */
	Integer setInteger(Number value);

	/**
	 * Set an integer at the current position
	 *
	 * @param value
	 *            the integer
	 * @param styleName
	 *            the name of the format
	 */
	Integer setInteger(Number value, String styleName);

	/**
	 * Set a format at the current position
	 *
	 * @param wrapperStyle
	 *            the format
	 */
	boolean setStyle(WrapperCellStyle wrapperStyle);

	/**
	 * Set a format at the current position
	 *
	 * @param styleName
	 *            the name of the format
	 */
	boolean setStyleName(String styleName);

	/**
	 * Set a text at the current position
	 *
	 * @param text
	 *            the text
	 */
	String setText(String text);

	/**
	 * Set a text at the current position
	 *
	 * @param text
	 *            the text
	 * @param styleName
	 *            the name of the format
	 */
	String setText(String text, String styleName);

}