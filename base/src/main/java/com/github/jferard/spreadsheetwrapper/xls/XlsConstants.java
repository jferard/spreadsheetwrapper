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
package com.github.jferard.spreadsheetwrapper.xls;

/**
 * Only a few constants for XLS.
 *
 */
public final class XlsConstants {
	/** standard extension for files */
	public static final String EXTENSION1 = "xls";

	/** extended extension for files */
	public static final String EXTENSION2 = "xlsx";

	/**
	 * The maximum number of columns per row
	 */
	public static final int MAX_COLUMNS = 256;

	/**
	 * The maximum number of rows excel allows in a worksheet
	 */
	public static final int MAX_ROWS_PER_SHEET = 65536;
	
	/** default font name */
	public static final String DEFAULT_FONT_NAME = "Arial"; 

	private XlsConstants() {
	}
}
