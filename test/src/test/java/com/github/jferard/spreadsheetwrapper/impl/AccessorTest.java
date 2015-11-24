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
import org.junit.Before;
import org.junit.Test;

public class AccessorTest {
	/** value for tests */
	private static final double FIRST_VALUE = 10.0;
	/** value for tests */
	private static final double SECOND_VALUE = 20.0;
	/** value for tests */
	private static final double THIRD_VALUE = 15.0;

	/** An accessor by name or index for double */
	private Accessor<Double> accessor;

	/** set the test up */
	@Before
	public void setUp() {
		this.accessor = new Accessor<Double>();
		this.accessor.putAtEnd("first", AccessorTest.FIRST_VALUE);
		this.accessor.putAtEnd("second", AccessorTest.SECOND_VALUE);
		this.accessor.put("firstb", 1, AccessorTest.THIRD_VALUE);
	}

	/** get Double by index */
	@Test
	public final void testGetByIndex() {
		Assert.assertEquals((Double) AccessorTest.FIRST_VALUE,
				this.accessor.getByIndex(0));
		Assert.assertEquals((Double) AccessorTest.THIRD_VALUE,
				this.accessor.getByIndex(1));
		Assert.assertEquals((Double) AccessorTest.SECOND_VALUE,
				this.accessor.getByIndex(2));
	}

	/** get Double by name */
	@Test
	public final void testGetByName() {
		Assert.assertEquals((Double) AccessorTest.FIRST_VALUE,
				this.accessor.getByName("first"));
		Assert.assertEquals((Double) AccessorTest.THIRD_VALUE,
				this.accessor.getByName("firstb"));
		Assert.assertEquals((Double) AccessorTest.SECOND_VALUE,
				this.accessor.getByName("second"));
	}

	/** has Double by index */
	@Test
	public final void testHasByIndex() {
		Assert.assertTrue(this.accessor.hasByIndex(0));
		Assert.assertTrue(this.accessor.hasByIndex(1));
		Assert.assertTrue(this.accessor.hasByIndex(2));
		Assert.assertFalse(this.accessor.hasByIndex(-1));
		Assert.assertFalse(this.accessor.hasByIndex(3));
	}

	/** has Double by name */
	@Test
	public final void testHasByName() {
		Assert.assertTrue(this.accessor.hasByName("first"));
		Assert.assertTrue(this.accessor.hasByName("firstb"));
		Assert.assertTrue(this.accessor.hasByName("second"));
		Assert.assertFalse(this.accessor.hasByName("third"));
	}

}
