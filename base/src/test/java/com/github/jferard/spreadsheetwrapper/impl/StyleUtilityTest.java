package com.github.jferard.spreadsheetwrapper.impl;

import org.junit.Assert;
import org.junit.Test;

import com.github.jferard.spreadsheetwrapper.StyleUtility;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;

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
				WrapperColor.GREY_25_PERCENT, WrapperCellStyle.DEFAULT, new WrapperFont(
						WrapperCellStyle.YES, WrapperCellStyle.YES, 15.0,  
						WrapperColor.DARK_BLUE, null));
		final String styleString0 = "background-color:GREY_25_PERCENT;font-weight:bold;font-style:italic;font-size:15.0;font-color:DARK_BLUE;";
		final String styleString1 = utility.toStyleString(cellStyle);
		Assert.assertEquals(styleString0, styleString1);
		final WrapperCellStyle cellStyle2 = utility
				.toWrapperCellStyle(styleString1);
		final String styleString2 = utility.toStyleString(cellStyle2);
		Assert.assertEquals(styleString1, styleString2);
	}

}
