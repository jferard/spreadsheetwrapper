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
 * The Cursor is on a cell of the spreadsheet.
 */
public interface Cursor { // NOPMD by Julien on 30/08/15 22:37
	/**
	 * NEXT : one cell move JUMP : several cells move FAIL : move fail (the
	 * position has not changed)
	 */
	public enum Move {
		FAIL, JUMP, NEXT
	}

	/**
	 * Moves to the first cell of the next row
	 *
	 * @return FAIL if was on the last row, JUMP else
	 */
	Move crlf();

	/**
	 * One row down
	 *
	 * @return NEXT if ok, FAIL if was on the last row
	 */
	Move down();

	/**
	 * Go to the first cell
	 *
	 * @return JUMP
	 */
	Move first();

	/**
	 * One row down
	 *
	 * @return NEXT
	 */
	Move forceDown();

	/**
	 * One column right
	 *
	 * @return NEXT
	 */
	Move forceRight();

	/**
	 * @return the column index (0..)
	 */
	int getC();

	/**
	 * @return the row index (0..)
	 */
	int getR();

	/**
	 * Go to the last cell
	 *
	 * @return JUMP
	 */
	Move last();

	/**
	 * One column left
	 *
	 * @return FAIL if was on column 0, NEXT else
	 */
	Move left();

	/**
	 * Moves to the next cell
	 *
	 * @return FAIL if was on the last cell, NEXT else
	 */
	Move next();

	/**
	 * Moves to the previous cell
	 *
	 * @return FAIL if was on the first cell, NEXT else
	 */
	Move previous();

	/**
	 * One column right
	 *
	 * @return FAIL if was on the last column, NEXT else
	 */
	Move right();

	/**
	 * @param r
	 *            the row : 0 .. n
	 * @param c
	 *            the column : 0 .. n
	 * @return the result
	 */
	Move set(final int r, final int c);

	/**
	 * One row up
	 *
	 * @return FAIL if was on the first row, JUMP else
	 */
	Move up(); // NOPMD by Julien on 30/08/15 22:37
}