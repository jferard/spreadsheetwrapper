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
 * A simple wrapper for some exceptions.
 */
public class CantInsertElementInSpreadsheetException extends
SpreadsheetException {
	/** for serialization */
	private static final long serialVersionUID = 1L;

	/**
	 * No message
	 */
	public CantInsertElementInSpreadsheetException() {
		super();
	}

	/**
	 * @param message
	 *            the message when there is nothing to wrap
	 */
	public CantInsertElementInSpreadsheetException(final String message) {
		super(message);
	}

	/**
	 * @param cause
	 *            the cause to be wrapped
	 */
	public CantInsertElementInSpreadsheetException(final Throwable cause) {
		super(cause);
	}
}
