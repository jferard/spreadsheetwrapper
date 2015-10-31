package com.github.jferard.spreadsheetwrapper.impl;

import org.junit.Assert;
import org.junit.Test;

import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;

public class StyleUtilityTest {

	@SuppressWarnings("static-method")
	@Test
	public final void test() {
		StyleUtility utility = new StyleUtility();
		final WrapperCellStyle cellStyle = new WrapperCellStyle(WrapperColor.GREY_25_PERCENT, new WrapperFont(WrapperCellStyle.YES, WrapperCellStyle.YES, 15, WrapperColor.DARK_BLUE));
		final String styleString0 = "background-color:GREY_25_PERCENT;font-weight:bold;font-style:italic;font-size:15;font-color:DARK_BLUE;";
		String styleString1 = utility.getStyleString(cellStyle);
		Assert.assertEquals(styleString0, styleString1);
		WrapperCellStyle cellStyle2 = utility.getWrapperCellStyle(styleString1);
		String styleString2 = utility.getStyleString(cellStyle2);
		Assert.assertEquals(styleString1, styleString2);
	}

}
