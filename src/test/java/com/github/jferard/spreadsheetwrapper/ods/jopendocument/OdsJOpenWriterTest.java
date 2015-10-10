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
package com.github.jferard.spreadsheetwrapper.ods.jopendocument;

import org.junit.Assert;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterTest;
import com.github.jferard.spreadsheetwrapper.TestProperties;

public class OdsJOpenWriterTest extends SpreadsheetWriterTest {
	@Override
	public final void testSetBoolean() {
		final int r = 5;
		final int c = 6;
		try {
			this.sw.setBoolean(r, c, true);
			Assert.assertEquals(true, this.sw.getBoolean(r, c));
		} catch (final IllegalArgumentException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		}
	}

	@Override
	protected TestProperties getProperties() {
		return OdsJOpenTestProperties.getProperties();
	}
}
