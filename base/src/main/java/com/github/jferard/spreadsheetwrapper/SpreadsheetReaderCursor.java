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

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public interface SpreadsheetReaderCursor extends Cursor {

	/**
	 * Get the content of the cell at the current position
	 *
	 * @return the content
	 */
	/*@Nullable*/Object getCellContent();

	/**
	 * Get the date at the current position
	 *
	 * @return the date
	 */
	Date getDate();

	/**
	 * Get the Double at the current position
	 *
	 * @return the Double
	 */
	Double getDouble();

	/**
	 * Get the formula at the current position
	 *
	 * @return the formula text (english)
	 */
	String getFormula();

	/**
	 * Get the integer at the current position
	 *
	 * @return the integer
	 */
	Integer getInteger();

	/**
	 * Get thee format at the current position
	 *
	 * @return the name of the format
	 */
	String getStyleName();

	/**
	 * Get thee format at the current position
	 *
	 * @return the name of the format
	 */
	String getStyleString();

	/**
	 * Get the text at the current position
	 *
	 * @return the text
	 */
	String getText();
}