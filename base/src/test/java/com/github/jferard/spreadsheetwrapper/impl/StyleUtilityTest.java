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

import com.github.jferard.spreadsheetwrapper.style.StyleUtility;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;

/**
 * Test class for style
 */
public class StyleUtilityTest {

	/** the single test */
	@SuppressWarnings("static-method")
	@Test
	public final void test() {
		final StyleUtility utility = new StyleUtility();
		final WrapperCellStyle cellStyle = new WrapperCellStyle(
				WrapperColor.GREY_25_PERCENT, WrapperCellStyle.DEFAULT,
				new WrapperFont(WrapperCellStyle.YES, WrapperCellStyle.YES,
						15.0, WrapperColor.DARK_BLUE, null));
		final String styleString0 = "background-color:GREY_25_PERCENT;font-weight:bold;font-style:italic;font-size:15.0;font-color:DARK_BLUE;";
		final String styleString1 = utility.toStyleString(cellStyle);
		Assert.assertEquals(styleString0, styleString1);
		final WrapperCellStyle cellStyle2 = utility
				.toWrapperCellStyle(styleString1);
		final String styleString2 = utility.toStyleString(cellStyle2);
		Assert.assertEquals(styleString1, styleString2);
	}

}
