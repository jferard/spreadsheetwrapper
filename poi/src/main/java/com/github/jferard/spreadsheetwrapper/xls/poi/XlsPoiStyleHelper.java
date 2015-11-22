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
import java.util.logging.Level;
import java.util.logging.Logger;

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
class XlsPoiStyleHelper {
	/** name or index -> cellStyle */
	private final CellStyleAccessor<CellStyle> cellStyleAccessor;
	/** internal -> wrapper */
	private final Map<HSSFColor, WrapperColor> colorByHssfColor;
	/** wrapper -> internal */
	private final Map<WrapperColor, HSSFColor> hssfColorByColor;

	/** simple logger */
	private final Logger logger;

	/**
	 * @param cellStyleAccessor
	 */
	public XlsPoiStyleHelper(
			final CellStyleAccessor<CellStyle> cellStyleAccessor) {
		this.cellStyleAccessor = cellStyleAccessor;
		this.logger = Logger.getLogger(this.getClass().getName());
		final WrapperColor[] colors = WrapperColor.values();
		this.hssfColorByColor = new HashMap<WrapperColor, HSSFColor>(
				colors.length);
		this.colorByHssfColor = new HashMap<HSSFColor, WrapperColor>(
				colors.length);
		final Map<Integer, HSSFColor> hssfColorByIndex = HSSFColor
				.getIndexHash();
		for (final HSSFColor hssfColor : hssfColorByIndex.values()) {
			final String hssfColorName = hssfColor.getClass().getName();
			final int index = hssfColorName.indexOf('$');
			if (index == -1)
				continue;

			final String colorName = hssfColorName.substring(index + 1);
			try {
				final WrapperColor wrapperColor = WrapperColor
						.valueOf(colorName);
				this.hssfColorByColor.put(wrapperColor, hssfColor);
				this.colorByHssfColor.put(hssfColor, wrapperColor);
			} catch (final IllegalArgumentException e) {
				this.logger
						.log(Level.WARNING,
								"Missing colors in WrapperColor class. Those colors won't be available for POI wrapper.",
								e);
			}
		}
	}

	/**
	 * @param workbook
	 *            internal workbook
	 * @param styleName
	 *            the name of the style
	 * @return the internal style, null if none
	 */
	public/*@Nullable*/CellStyle getCellStyle(final Workbook workbook,
			final String styleName) {
		CellStyle cellStyle = this.cellStyleAccessor.getCellStyle(styleName);
		if (cellStyle == null && styleName.startsWith("ssw")) {
			try {
				final int idx = Integer.valueOf(styleName.substring(3));
				cellStyle = workbook.getCellStyleAt((short) idx);
			} catch (final NumberFormatException e) { // NOPMD by Julien on
														// 22/11/15 07:15
				// do nothing
			}
		}
		return cellStyle;
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
			final HSSFColor hssfColor = this.toHSSFColor(backgroundColor);

			final short index = hssfColor.getIndex();
			cellStyle.setFillForegroundColor(index);
			cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		}
		return cellStyle;
	}

	/**
	 * @param cellStyle
	 *            the internal style
	 * @return the style name, ssw<index> if none.
	 */
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
		final StringBuilder styleStringBuilder = new StringBuilder(25);
		final short fontIndex = cellStyle.getFontIndex();
		final Font font = workbook.getFontAt(fontIndex);
		if (font.getBoldweight() == Font.BOLDWEIGHT_BOLD)
			styleStringBuilder.append("font-weight:bold;");
		final Color color = cellStyle.getFillBackgroundColorColor();
		if (color instanceof HSSFColor) // color != null
			styleStringBuilder.append("background-color:")
			.append(((HSSFColor) color).getHexString()).append(';');
		return styleStringBuilder.toString();
	}

	/**
	 * Create or update style
	 * 
	 * @param styleName
	 *            the name of the style
	 * @param cellStyle
	 *            the internal style to put
	 */
	public void putCellStyle(final String styleName, final CellStyle cellStyle) {
		this.cellStyleAccessor.putCellStyle(styleName, cellStyle);
	}

	/**
	 * @param wrapperColor
	 *            the color to convert
	 * @return the HSSF color
	 */
	public HSSFColor toHSSFColor(final WrapperColor wrapperColor) {
		HSSFColor hssfColor;
		if (this.hssfColorByColor.containsKey(wrapperColor))
			hssfColor = this.hssfColorByColor.get(wrapperColor);
		else
			hssfColor = new HSSFColor.WHITE();
		return hssfColor;
	}

	/**
	 * @param workbook
	 *            the *internal* workbook
	 * @param cellStyle
	 *            the style in poi format
	 * @return the cell style in new format
	 */
	public WrapperCellStyle toWrapperCellStyle(final Workbook workbook,
			final CellStyle cellStyle) {
		if (cellStyle == null)
			return WrapperCellStyle.EMPTY;

		final short fontIndex = cellStyle.getFontIndex();
		final Font poiFont = workbook.getFontAt(fontIndex);
		WrapperCellStyle wrapperCellStyle;
		WrapperFont wrapperFont;
		if (poiFont.getBoldweight() == Font.BOLDWEIGHT_BOLD)
			wrapperFont = new WrapperFont().setBold();
		else
			wrapperFont = new WrapperFont();

		WrapperColor wrapperColor = null;
		final short index = cellStyle.getFillForegroundColor();
		final Map<Integer, HSSFColor> indexHash = HSSFColor.getIndexHash();
		final HSSFColor poiColor = indexHash.get(Integer.valueOf(index));
		if (cellStyle.getFillPattern() == CellStyle.SOLID_FOREGROUND
				&& this.colorByHssfColor.containsKey(poiColor))
			wrapperColor = this.colorByHssfColor.get(poiColor);
		if (WrapperColor.DEFAULT_BACKGROUND.equals(wrapperColor))
			wrapperColor = null;

		wrapperCellStyle = new WrapperCellStyle(wrapperColor, wrapperFont);
		return wrapperCellStyle;
	}
}
