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
	private Accessor<Double> accessor;

	@Before
	public void setUp() {
		this.accessor = new Accessor<Double>();
		this.accessor.putAtEnd("first", 10.0);
		this.accessor.putAtEnd("second", 20.0);
		this.accessor.put("firstb", 1, 15.0);
	}

	@Test
	public final void testIndex() {
		Assert.assertTrue(this.accessor.hasByIndex(0));
		Assert.assertTrue(this.accessor.hasByIndex(1));
		Assert.assertTrue(this.accessor.hasByIndex(2));
		Assert.assertFalse(this.accessor.hasByIndex(-1));
		Assert.assertFalse(this.accessor.hasByIndex(3));
	}

	@Test
	public final void testIndex2() {
		Assert.assertEquals((Double) 10.0, this.accessor.getByIndex(0));
		Assert.assertEquals((Double) 15.0, this.accessor.getByIndex(1));
		Assert.assertEquals((Double) 20.0, this.accessor.getByIndex(2));
	}

	@Test
	public final void testName() {
		Assert.assertTrue(this.accessor.hasByName("first"));
		Assert.assertTrue(this.accessor.hasByName("firstb"));
		Assert.assertTrue(this.accessor.hasByName("second"));
		Assert.assertFalse(this.accessor.hasByName("third"));
	}

	@Test
	public final void testName2() {
		Assert.assertEquals((Double) 10.0, this.accessor.getByName("first"));
		Assert.assertEquals((Double) 15.0, this.accessor.getByName("firstb"));
		Assert.assertEquals((Double) 20.0, this.accessor.getByName("second"));
	}

}
