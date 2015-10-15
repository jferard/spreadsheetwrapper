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

/**
 * Visitor pattern. A wrapper for data to write on the sheet One should declare
 * a wrapper an the method to write it on the sheet E.g. a ResultSetDataWrapper
 * has to declare the writeDataTo method that writes the content of the
 * resultSet on the sheet.
 */
public interface DataWrapper {
	/**
	 * @param writer
	 *            sheet where to write
	 * @param r
	 *            row index
	 * @param c
	 *            column inex
	 * @return true if the data is written
	 */
	boolean writeDataTo(SpreadsheetWriter writer, int r, int c);
}
