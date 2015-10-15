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
package com.github.jferard.spreadsheetwrapper.impl;

import com.github.jferard.spreadsheetwrapper.Cursor;

/**
 */
public abstract class AbstractCursor implements Cursor { // NOPMD by Julien on
	// 03/09/15 22:11
	/** the current column */
	private int c;
	/** the current row */
	private int r;
	/** the row number */
	private final int rowCount;

	/**
	 * @param rowCount
	 *            the numbe of row of the sheet, table, ...
	 */
	public AbstractCursor(final int rowCount) {
		this.rowCount = rowCount;
		this.r = 0;
		this.c = 0;
	}

	/** {@inheritDoc} */
	@Override
	public Move crlf() {
		if (this.r >= this.rowCount - 1)
			return Move.FAIL;

		this.c = 0;
		this.r++;
		return Move.JUMP;
	}

	/** {@inheritDoc} */
	@Override
	public Move down() {
		if (this.r >= this.rowCount - 1)
			return Move.FAIL;

		this.r++;
		return Move.NEXT;
	}

	/** {@inheritDoc} */
	@Override
	public Move first() {
		this.r = 0;
		this.c = 0;
		return Move.JUMP;
	}

	/** {@inheritDoc} */
	@Override
	public Move forceDown() {
		this.r++;
		return Move.NEXT;
	}

	/** {@inheritDoc} */
	@Override
	public Move forceRight() {
		this.c++;
		return Move.NEXT;
	}

	/** {@inheritDoc} */
	@Override
	public int getC() {
		return this.c;
	}

	/** {@inheritDoc} */
	@Override
	public int getR() {
		return this.r;
	}

	/** {@inheritDoc} */
	@Override
	public Move last() {
		this.r = this.rowCount - 1;
		this.c = this.getCellCount(this.r) - 1;
		return Move.JUMP;
	}

	/** {@inheritDoc} */
	@Override
	public Move left() {
		if (this.c == 0)
			return Move.FAIL;

		this.c--;
		return Move.NEXT;
	}

	/** {@inheritDoc} */
	@Override
	public Move next() {
		Move move;
		if (this.c >= this.getCellCount(this.r) - 1) {
			if (this.r >= this.rowCount - 1)
				move = Move.FAIL;
			else {
				this.c = 0;
				this.r++;
				move = Move.JUMP;
			}
		} else {
			this.c++;
			move = Move.NEXT;
		}
		return move;
	}

	/** {@inheritDoc} */
	@Override
	public Move previous() {
		Move move;
		if (this.c == 0) {
			if (this.r == 0)
				move = Move.FAIL;
			else {
				this.r--;
				this.c = this.getCellCount(this.r) - 1;
				move = Move.JUMP;
			}
		} else {
			this.c--;
			move = Move.NEXT;
		}
		return move;
	}

	/** {@inheritDoc} */
	@Override
	public Move right() {
		if (this.c >= this.getCellCount(this.r) - 1)
			return Move.FAIL;

		this.c++;
		return Move.NEXT;
	}

	/** {@inheritDoc} */
	@Override
	public Move set(final int newR, final int newC) {
		if (newR < 0 || newR >= this.rowCount || newC < 0
				|| newC >= this.getCellCount(newR))
			return Move.FAIL;

		this.c = newC;
		this.r = newR;
		return Move.JUMP;
	}

	/** {@inheritDoc} */
	@Override
	public Move up() { // NOPMD by Julien on 03/09/15 22:11
		if (this.r == 0)
			return Move.FAIL;

		this.r--;
		return Move.NEXT;
	}

	/**
	 * @param r
	 *            row index (0..)
	 * @return the number of cells in the row
	 */
	protected abstract int getCellCount(int r);
}