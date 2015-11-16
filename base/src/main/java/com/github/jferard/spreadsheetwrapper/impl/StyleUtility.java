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

import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;

public class StyleUtility {
	public static final String BACKGROUND_COLOR = "background-color";
	public static final String FONT_COLOR = "font-color";
	public static final String FONT_SIZE = "font-size";
	public static final String FONT_STYLE = "font-style";
	public static final String FONT_WEIGHT = "font-weight";

	public String getStyleString(final WrapperCellStyle cellStyle) {
		final StringBuilder styleStringBuilder = new StringBuilder();
		final WrapperColor backgroundColor = cellStyle.getBackgroundColor();
		if (backgroundColor != null)
			styleStringBuilder.append(StyleUtility.BACKGROUND_COLOR)
					.append(':').append(backgroundColor.name()).append(';');
		final WrapperFont font = cellStyle.getCellFont();
		final int size;
		final WrapperColor color;
		if (font == null) {
			size = WrapperCellStyle.DEFAULT;
			color = null;
		} else {
			size = font.getSize();
			color = font.getColor();
			switch (font.getBold()) {
			case WrapperCellStyle.YES:
				styleStringBuilder.append(StyleUtility.FONT_WEIGHT).append(
						":bold;");
				break;
			case WrapperCellStyle.NO:
				styleStringBuilder.append(StyleUtility.FONT_WEIGHT).append(
						":normal;");
				break;
			default:
				break;

			}
			switch (font.getItalic()) {
			case WrapperCellStyle.YES:
				styleStringBuilder.append(StyleUtility.FONT_STYLE).append(
						":italic;");
				break;
			case WrapperCellStyle.NO:
				styleStringBuilder.append(StyleUtility.FONT_STYLE).append(
						":normal;");
				break;
			default:
				break;
			}
		}
		if (size != WrapperCellStyle.DEFAULT)
			styleStringBuilder.append(StyleUtility.FONT_SIZE).append(':')
					.append(size).append(';');
		if (color != null)
			styleStringBuilder.append(StyleUtility.FONT_COLOR).append(':')
					.append(color.name()).append(';');

		return styleStringBuilder.toString();
	}

	public WrapperCellStyle getWrapperCellStyle(final String styleString) {
		final String[] styleProps = styleString.split(";");
		WrapperColor backgroundColor = null;
		int bold = WrapperCellStyle.DEFAULT;
		int italic = WrapperCellStyle.DEFAULT;
		int size = WrapperCellStyle.DEFAULT;
		WrapperColor color = null;
		for (final String styleProp : styleProps) {
			final String[] entry = styleProp.split(":");
			if (entry.length != 2)
				throw new IllegalArgumentException(styleString);

			final String key = entry[0];
			final String value = entry[1];
			if (key.equals(StyleUtility.BACKGROUND_COLOR))
				backgroundColor = WrapperColor.getColorFromString(value);
			else if (key.equals(StyleUtility.FONT_WEIGHT)) {
				if (value.equals("bold"))
					bold = WrapperCellStyle.YES;
				else if (value.equals("normal"))
					bold = WrapperCellStyle.NO;
			} else if (key.equals(StyleUtility.FONT_STYLE)) {
				if (value.equals("italic"))
					italic = WrapperCellStyle.YES;
				else if (value.equals("normal"))
					italic = WrapperCellStyle.NO;
			} else if (key.equals(StyleUtility.FONT_SIZE))
				size = Integer.valueOf(value);
			else if (key.equals(StyleUtility.FONT_COLOR))
				color = WrapperColor.getColorFromString(value);
		}

		return new WrapperCellStyle(backgroundColor, new WrapperFont(bold,
				italic, size, color));
	}
}
