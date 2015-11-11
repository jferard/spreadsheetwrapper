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

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;
import com.github.jferard.spreadsheetwrapper.impl.CellStyleAccessor;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * An utility class for style handling
 *
 */
public class XlsPoiStyleHelper {
	private final CellStyleAccessor<CellStyle> cellStyleAccessor;
	private final Map<HSSFColor, WrapperColor> colorByHssfColor;
	private final Map<WrapperColor, HSSFColor> hssfColorByColor;

	/**
	 * @param cellStyleAccessor
	 */
	public XlsPoiStyleHelper(
			final CellStyleAccessor<CellStyle> cellStyleAccessor) {
		this.cellStyleAccessor = cellStyleAccessor;
		final WrapperColor[] colors = WrapperColor.values();
		this.hssfColorByColor = new HashMap<WrapperColor, HSSFColor>(
				colors.length);
		this.colorByHssfColor = new HashMap<HSSFColor, WrapperColor>(
				colors.length);
		for (final WrapperColor color : colors) {
			HSSFColor hssfColor;
			try {
				final Class<?> hssfClazz = Class
						.forName("org.apache.poi.hssf.util.HSSFColor$"
								+ color.getSimpleName());
				hssfColor = (HSSFColor) hssfClazz.newInstance();
				this.hssfColorByColor.put(color, hssfColor);
				this.colorByHssfColor.put(hssfColor, color);
			} catch (final ClassNotFoundException e) {
				hssfColor = null;
			} catch (final IllegalArgumentException e) {
				hssfColor = null;
			} catch (final IllegalAccessException e) {
				hssfColor = null;
			} catch (final SecurityException e) {
				hssfColor = null;
			} catch (final InstantiationException e) {
				hssfColor = null;
			}
		}
	}

	public CellStyle getCellStyle(final Workbook workbook,
			final WrapperCellStyle wrapperCellStyle) {
		final CellStyle cellStyle = workbook.createCellStyle();
		final WrapperFont wrapperFont = wrapperCellStyle.getCellFont();
		if (wrapperFont != null
				&& wrapperFont.getBold() == WrapperCellStyle.YES) {
			final Font font = workbook.createFont();
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			cellStyle.setFont(font);
		}
		final WrapperColor backgroundColor = wrapperCellStyle
				.getBackgroundColor();
		if (backgroundColor != null) {
			final HSSFColor hssfColor = this.getHSSFColor(backgroundColor);

			final short index = hssfColor.getIndex();
			cellStyle.setFillForegroundColor(index);
			cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		}
		return cellStyle;
	}

	public HSSFColor getHSSFColor(final WrapperColor backgroundColor) {
		HSSFColor hssfColor;
		if (this.hssfColorByColor.containsKey(backgroundColor))
			hssfColor = this.hssfColorByColor.get(backgroundColor);
		else
			hssfColor = new HSSFColor.WHITE();
		return hssfColor;
	}

	public String getStyleName(final CellStyle cellStyle) {
		final String name = this.cellStyleAccessor.getName(cellStyle);
		if (name == null)
			return "ssw" + cellStyle.getIndex();
		else
			return name;
	}

	/**
	 * Reverse of the getCellStyle method
	 *
	 * @param workbook
	 *            the *internal* workbook
	 * @param cellStyle
	 *            the style in poi format
	 * @return the old style string
	 */
	public String getStyleString(final Workbook workbook,
			final CellStyle cellStyle) {
		final StringBuilder sb = new StringBuilder();
		final short fontIndex = cellStyle.getFontIndex();
		final Font font = workbook.getFontAt(fontIndex);
		if (font.getBoldweight() == Font.BOLDWEIGHT_BOLD)
			sb.append("font-weight:bold;");
		final Color color = cellStyle.getFillBackgroundColorColor();
		if (color != null && color instanceof HSSFColor)
			sb.append("background-color:")
			.append(((HSSFColor) color).getHexString()).append(";");
		return sb.toString();
	}

	public void putCellStyle(final String styleName, final CellStyle cellStyle) {
		this.cellStyleAccessor.putCellStyle(styleName, cellStyle);
	}

	CellStyle getCellStyle(final Workbook workbook, final String styleName) {
		CellStyle cellStyle = this.cellStyleAccessor.getCellStyle(styleName);
		if (cellStyle == null) {
			if (styleName.startsWith("ssw")) {
				try {
					final int idx = Integer.valueOf(styleName.substring(3));
					cellStyle = workbook.getCellStyleAt((short) idx);
				} catch (final NumberFormatException e) {
					// do nothing
				}
			}
		}
		return cellStyle;
	}

	/**
	 * @param workbook
	 *            the *internal* workbook
	 * @param cellStyle
	 *            the style in poi format
	 * @return the cell style in new format
	 */
	WrapperCellStyle getWrapperCellStyle(final Workbook workbook,
			final CellStyle cellStyle) {
		final short fontIndex = cellStyle.getFontIndex();
		final Font poiFont = workbook.getFontAt(fontIndex);
		WrapperCellStyle wrapperCellStyle;
		WrapperFont wrapperFont;
		if (poiFont.getBoldweight() == Font.BOLDWEIGHT_BOLD)
			wrapperFont = new WrapperFont().setBold();
		else
			wrapperFont = new WrapperFont();

		final Color poiColor = cellStyle.getFillBackgroundColorColor();
		final WrapperColor wrapperColor;
		if (this.colorByHssfColor.containsKey(poiColor))
			wrapperColor = this.colorByHssfColor.get(poiColor);
		else
			wrapperColor = WrapperColor.AUTOMATIC;

		wrapperCellStyle = new WrapperCellStyle(wrapperColor, wrapperFont);
		return wrapperCellStyle;
	}
}
