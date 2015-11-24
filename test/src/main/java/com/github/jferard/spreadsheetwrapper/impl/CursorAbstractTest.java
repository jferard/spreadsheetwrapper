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

import org.junit.Assert;
import org.junit.Test;

import com.github.jferard.spreadsheetwrapper.Cursor;
import com.github.jferard.spreadsheetwrapper.Cursor.Move;

public class CursorAbstractTest {

	/** the column count */
	protected int colCount;

	/** the cursor */
	protected Cursor cursor;

	/** the row count */
	protected int rowCount;

	/** A1 + left = fail, A1 + up = fail, A1 + right = B1, A1 + down = A2 */
	@Test
	public final void testCantGoLeftOrUpFromFirstCell() {
		Assert.assertEquals(Move.JUMP, this.cursor.first());
		Assert.assertEquals(0, this.cursor.getR());
		Assert.assertEquals(0, this.cursor.getC());
		Assert.assertEquals(Move.FAIL, this.cursor.left());
		Assert.assertEquals(0, this.cursor.getR());
		Assert.assertEquals(0, this.cursor.getC());
		Assert.assertEquals(Move.FAIL, this.cursor.up());
		Assert.assertEquals(0, this.cursor.getR());
		Assert.assertEquals(0, this.cursor.getC());
		Assert.assertEquals(Move.NEXT, this.cursor.right());
		Assert.assertEquals(0, this.cursor.getR());
		Assert.assertEquals(1, this.cursor.getC());
		Assert.assertEquals(Move.NEXT, this.cursor.down());
		Assert.assertEquals(1, this.cursor.getR());
		Assert.assertEquals(1, this.cursor.getC());
	}

	/** RC + right = fail, RC + down = fail, RC + left= RC-1, RC + up = R-1C */
	@Test
	public final void testCantGoRightOrownFromLastCell() {
		Assert.assertEquals(Move.JUMP, this.cursor.last());
		Assert.assertEquals(this.rowCount - 1, this.cursor.getR());
		Assert.assertEquals(this.colCount - 1, this.cursor.getC());
		Assert.assertEquals(Move.FAIL, this.cursor.right());
		Assert.assertEquals(this.rowCount - 1, this.cursor.getR());
		Assert.assertEquals(this.colCount - 1, this.cursor.getC());
		Assert.assertEquals(Move.FAIL, this.cursor.down());
		Assert.assertEquals(this.rowCount - 1, this.cursor.getR());
		Assert.assertEquals(this.colCount - 1, this.cursor.getC());
		Assert.assertEquals(Move.NEXT, this.cursor.left());
		Assert.assertEquals(this.rowCount - 1, this.cursor.getR());
		Assert.assertEquals(this.colCount - 2, this.cursor.getC());
		Assert.assertEquals(Move.NEXT, this.cursor.up());
		Assert.assertEquals(this.rowCount - 2, this.cursor.getR());
		Assert.assertEquals(this.colCount - 2, this.cursor.getC());
	}

	/** A1 + next + next + previous = B1 */
	@Test
	public final void testFirstIsA1NextIsA2() {
		Assert.assertEquals(Move.JUMP, this.cursor.first());
		Assert.assertEquals(0, this.cursor.getR());
		Assert.assertEquals(0, this.cursor.getC());
		Assert.assertEquals(Move.NEXT, this.cursor.next());
		Assert.assertEquals(Move.NEXT, this.cursor.next());
		Assert.assertEquals(Move.NEXT, this.cursor.previous());
		Assert.assertEquals(0, this.cursor.getR());
		Assert.assertEquals(1, this.cursor.getC());
	}

	/** cursor.set + various incorrect values */
	@Test
	public final void testIncorrectCoordinatesFail() {
		Assert.assertEquals(Move.FAIL, this.cursor.set(-1, 2));
		Assert.assertEquals(Move.FAIL, this.cursor.set(1, -2));
		Assert.assertEquals(Move.FAIL, this.cursor.set(this.rowCount * 2, 2));
		Assert.assertEquals(Move.FAIL, this.cursor.set(1, this.colCount * 2));
	}

	/** RC + crlf = fail */
	@Test
	public final void testNoCRLFFromLastCell() {
		Assert.assertEquals(Move.JUMP, this.cursor.last());
		Assert.assertEquals(this.rowCount - 1, this.cursor.getR());
		Assert.assertEquals(this.colCount - 1, this.cursor.getC());
		Assert.assertEquals(Move.FAIL, this.cursor.crlf());
		Assert.assertEquals(this.rowCount - 1, this.cursor.getR());
		Assert.assertEquals(this.colCount - 1, this.cursor.getC()); // no reset
		// of the
		// col
	}

	/** A1 + previous = fail */
	@Test
	public final void testNoPreviousFromlFirstCell() {
		Assert.assertEquals(Move.JUMP, this.cursor.first());
		Assert.assertEquals(0, this.cursor.getR());
		Assert.assertEquals(0, this.cursor.getC());
		Assert.assertEquals(Move.FAIL, this.cursor.previous());
		Assert.assertEquals(0, this.cursor.getR());
		Assert.assertEquals(0, this.cursor.getC()); // no reset of the col
	}

	/** a rotation around a cell */
	@Test
	public final void testRotation() {
		final int r = this.rowCount / 2;
		final int c = this.colCount / 2;
		Assert.assertEquals(Move.JUMP, this.cursor.set(r, c));
		Assert.assertEquals(r, this.cursor.getR());
		Assert.assertEquals(c, this.cursor.getC());
		this.cursor.right();
		Assert.assertEquals(r, this.cursor.getR());
		Assert.assertEquals(c + 1, this.cursor.getC());
		this.cursor.up();
		Assert.assertEquals(r - 1, this.cursor.getR());
		Assert.assertEquals(c + 1, this.cursor.getC());
		this.cursor.left();
		Assert.assertEquals(r - 1, this.cursor.getR());
		Assert.assertEquals(c, this.cursor.getC());
		this.cursor.left();
		Assert.assertEquals(r - 1, this.cursor.getR());
		Assert.assertEquals(c - 1, this.cursor.getC());
		this.cursor.down();
		Assert.assertEquals(r, this.cursor.getR());
		Assert.assertEquals(c - 1, this.cursor.getC());
		this.cursor.right();
		Assert.assertEquals(r, this.cursor.getR());
		Assert.assertEquals(c, this.cursor.getC());
	}

	/** a line jump with next */
	@Test
	public final void testRowEnd() {
		Assert.assertEquals(Move.JUMP, this.cursor.set(0, this.colCount - 1));
		Assert.assertEquals(0, this.cursor.getR());
		Assert.assertEquals(this.colCount - 1, this.cursor.getC());
		Assert.assertEquals(Move.JUMP, this.cursor.next()); // because of the
		// line jump
		Assert.assertEquals(1, this.cursor.getR());
		Assert.assertEquals(0, this.cursor.getC());
		Assert.assertEquals(Move.JUMP, this.cursor.previous()); // because of
		// the line jump
		Assert.assertEquals(0, this.cursor.getR());
		Assert.assertEquals(this.colCount - 1, this.cursor.getC());
		Assert.assertEquals(Move.NEXT, this.cursor.previous());
		Assert.assertEquals(Move.JUMP, this.cursor.crlf());
		Assert.assertEquals(1, this.cursor.getR());
		Assert.assertEquals(0, this.cursor.getC());
	}
}