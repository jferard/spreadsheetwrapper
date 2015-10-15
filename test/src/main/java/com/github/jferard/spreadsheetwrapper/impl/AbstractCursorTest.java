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

import org.junit.Before;

/**
 * An abstract test for cursor
 *
 */
public class AbstractCursorTest extends CursorAbstractTest {
	/**
	 * Sets the test up : a 20*20 table, with a basic cursor.
	 */
	@Before
	public void setUp() {
		this.rowCount = 19;
		this.colCount = 19;
		this.cursor = new AbstractCursor(this.rowCount) {
			@Override
			protected int getCellCount(final int r) {
				return 19;
			}
		};
	}
}
