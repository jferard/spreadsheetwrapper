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
package com.github.jferard.spreadsheetwrapper.xls.poi;

import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyleHelper;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;
import com.github.jferard.spreadsheetwrapper.impl.StyleUtility;

/**
 * An utility class for style handling
 *
 */
public class XlsPoiStyleUtility extends StyleUtility {
	/** for colors */
	private final WrapperCellStyleHelper helper;

	/**
	 * @param helper
	 *            helper class for colors
	 */
	public XlsPoiStyleUtility(final WrapperCellStyleHelper helper) {
		this.helper = helper;
	}

	public CellStyle getCellStyle(final Workbook workbook,
			final String styleString) {
		final CellStyle cellStyle = workbook.createCellStyle();
		final Map<String, String> props = this.getPropertiesMap(styleString);
		for (final Map.Entry<String, String> entry : props.entrySet()) {
			if (entry.getKey().equals("font-weight")) {
				if (entry.getValue().equals("bold")) {
					final Font font = workbook.createFont();
					font.setBoldweight(Font.BOLDWEIGHT_BOLD);
					cellStyle.setFont(font);
				}
			} else if (entry.getKey().equals("background-color")) {
				// do nothing
			}
		}
		return cellStyle;
	}

	public String getStyleString(final Workbook workbook,
			final CellStyle cellStyle) {
		final StringBuilder sb = new StringBuilder();
		final short fontIndex = cellStyle.getFontIndex();
		final Font font = workbook.getFontAt(fontIndex);
		if (font.getBoldweight() == Font.BOLDWEIGHT_BOLD)
			sb.append("font-weight:bold;");
		final Color color = cellStyle.getFillBackgroundColorColor();
		return sb.toString();
	}

	WrapperCellStyle getCellStyle(final Workbook workbook,
			final CellStyle cellStyle) {
		final short fontIndex = cellStyle.getFontIndex();
		final Font poiFont = workbook.getFontAt(fontIndex);
		WrapperCellStyle wrapperCellStyle = null;
		WrapperFont wrapperFont = null;
		if (poiFont.getBoldweight() == Font.BOLDWEIGHT_BOLD)
			wrapperFont = new WrapperFont(true, false, fontIndex,
					WrapperColor.AUTOMATIC);
		final Color poiColor = cellStyle.getFillBackgroundColorColor();
		WrapperColor wrapperColor = this.helper.getColor(poiColor);

		wrapperCellStyle = new WrapperCellStyle(wrapperColor, wrapperFont);
		return wrapperCellStyle;
	}
}
