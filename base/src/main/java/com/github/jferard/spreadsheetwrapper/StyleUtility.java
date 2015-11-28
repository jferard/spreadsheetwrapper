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
package com.github.jferard.spreadsheetwrapper;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URL;
import java.net.URLConnection;
import java.net.UnknownServiceException;

public class StyleUtility {

	/** for style string */
	public static final String BACKGROUND_COLOR = "background-color";
	/** for style string */
	public static final String FONT_COLOR = "font-color";
	/** for style string */
	public static final String FONT_SIZE = "font-size";
	/** for style string */
	public static final String FONT_STYLE = "font-style";
	/** for style string */
	public static final String FONT_WEIGHT = "font-weight";

	/**
	 * @param outputURL
	 *            the URL of the file to write
	 * @return the stream on the file
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public static OutputStream getOutputStream(final URL outputURL)
			throws IOException, FileNotFoundException {
		OutputStream outputStream;
		final URLConnection connection = outputURL.openConnection();
		connection.setDoOutput(true);
		try {
			outputStream = connection.getOutputStream();
		} catch (final UnknownServiceException e) {
			outputStream = new FileOutputStream(outputURL.getPath());
		}
		return outputStream;
	}

	private static void appendFontBold(final StringBuilder styleStringBuilder,
			final int bold) {
		switch (bold) {
		case WrapperCellStyle.YES:
			styleStringBuilder.append(StyleUtility.FONT_WEIGHT)
			.append(":bold;");
			break;
		case WrapperCellStyle.NO:
			styleStringBuilder.append(StyleUtility.FONT_WEIGHT).append(
					":normal;");
			break;
		default:
			break;

		}
	}

	private static void appendFontItalic(
			final StringBuilder styleStringBuilder, final int italic) {
		switch (italic) {
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

	private static void setFontBold(final WrapperFont font, final String value) {
		if (value.equals("bold"))
			font.setBold();
		else if (value.equals("normal"))
			font.setBold(WrapperCellStyle.NO);
	}

	private static void setFontItalic(final WrapperFont font, final String value) {
		if (value.equals("italic"))
			font.setItalic();
		else if (value.equals("normal"))
			font.setItalic(WrapperCellStyle.NO);
	}

	/**
	 * Convert a WrapperCellStyle to a styleString (CSS-like format)
	 *
	 * @param cellStyle
	 *            the cell style
	 * @return the string to convert
	 */
	/**
	 * @param cellStyle
	 * @return
	 */
	public String toStyleString(final WrapperCellStyle cellStyle) {
		final StringBuilder styleStringBuilder = new StringBuilder(20);
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
			StyleUtility.appendFontBold(styleStringBuilder, font.getBold());
			StyleUtility.appendFontItalic(styleStringBuilder, font.getItalic());
		}
		if (size != WrapperCellStyle.DEFAULT)
			styleStringBuilder.append(StyleUtility.FONT_SIZE).append(':')
					.append(size).append(';');
		if (color != null)
			styleStringBuilder.append(StyleUtility.FONT_COLOR).append(':')
					.append(color.name()).append(';');

		return styleStringBuilder.toString();
	}

	/**
	 * Convert a styleString (CSS-like format) to a WrapperCellStyle
	 *
	 * @param styleString
	 *            the string to convert
	 * @return the cell style
	 */
	public WrapperCellStyle toWrapperCellStyle(final String styleString) {
		final String[] styleProps = styleString.split(";");
		WrapperColor backgroundColor = null;
		final WrapperFont font = new WrapperFont();
		for (final String styleProp : styleProps) {
			final String[] entry = styleProp.split(":");
			if (entry.length != 2) // NOPMD by Julien on 21/11/15 10:38
				throw new IllegalArgumentException(styleString);

			final String key = entry[0];
			final String value = entry[1];
			if (key.equals(StyleUtility.BACKGROUND_COLOR))
				backgroundColor = WrapperColor.stringToColor(value);
			else if (key.equals(StyleUtility.FONT_WEIGHT))
				StyleUtility.setFontBold(font, value);
			else if (key.equals(StyleUtility.FONT_STYLE))
				StyleUtility.setFontItalic(font, value);
			else if (key.equals(StyleUtility.FONT_SIZE))
				font.setSize(Integer.valueOf(value));
			else if (key.equals(StyleUtility.FONT_COLOR))
				font.setColor(WrapperColor.stringToColor(value));
		}

		return new WrapperCellStyle(backgroundColor, font);
	}

}
