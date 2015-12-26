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

import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.jferard.spreadsheetwrapper.Util;
import com.github.jferard.spreadsheetwrapper.style.Borders;
import com.github.jferard.spreadsheetwrapper.style.CellStyleAccessor;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;
import com.github.jferard.spreadsheetwrapper.xls.XlsConstants;

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

	/**
	 * @param cellStyleAccessor
	 */
	public XlsPoiStyleHelper(
			final CellStyleAccessor<CellStyle> cellStyleAccessor) {
		this.cellStyleAccessor = cellStyleAccessor;
		final Logger logger = Logger.getLogger(this.getClass().getName());
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
				logger.log(
						Level.WARNING,
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
	 * @param workbook
	 *            workbook for conversion
	 * @param wrapperCellStyle
	 *            the cell style
	 * @return the internal cell style
	 */
	public CellStyle toCellStyle(final Workbook workbook,
			final WrapperCellStyle wrapperCellStyle) {
		final CellStyle cellStyle = workbook.createCellStyle();
		final WrapperFont wrapperFont = wrapperCellStyle.getCellFont();
		if (wrapperFont != null) {
			final Font font = workbook.createFont();

			final int bold = wrapperFont.getBold();
			final int italic = wrapperFont.getItalic();
			final double sizeInPoints = wrapperFont.getSize();
			final WrapperColor fontColor = wrapperFont.getColor();
			final String fontFamily = wrapperFont.getFamily();

			if (bold == WrapperCellStyle.YES)
				font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			else if (bold == WrapperCellStyle.NO)
				font.setBoldweight(Font.BOLDWEIGHT_NORMAL);

			if (italic == WrapperCellStyle.YES)
				font.setItalic(true);
			else if (italic == WrapperCellStyle.NO)
				font.setItalic(false);

			if (!Util.almostEqual(sizeInPoints, WrapperCellStyle.DEFAULT)) {
				final double sizeInTwentiethOfPoint = sizeInPoints * 20.0;
				font.setFontHeight((short) sizeInTwentiethOfPoint);
			}

			if (fontColor != null) {
				final HSSFColor hssfColor = this.toHSSFColor(fontColor);
				final short index = hssfColor.getIndex();
				font.setColor(index);
			}

			if (fontFamily != null)
				font.setFontName(fontFamily);

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

		
		final /*@Nullable*/ Borders borders = wrapperCellStyle.getBorders();
		if (borders != null) {
			final double borderWidth = borders.getLineWidth();
			if (borderWidth != WrapperCellStyle.DEFAULT) {
				if (Util.almostEqual(borderWidth, WrapperCellStyle.THIN_LINE))
					this.setAllborders(cellStyle, CellStyle.BORDER_THIN);
				else if (Util
						.almostEqual(borderWidth, WrapperCellStyle.MEDIUM_LINE))
					this.setAllborders(cellStyle, CellStyle.BORDER_MEDIUM);
				else if (Util.almostEqual(borderWidth, WrapperCellStyle.THICK_LINE))
					this.setAllborders(cellStyle, CellStyle.BORDER_THICK);
				else
					throw new UnsupportedOperationException();
			}
		}
		return cellStyle;
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
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle();
		final WrapperFont wrapperFont = new WrapperFont();
		if (poiFont.getBoldweight() == Font.BOLDWEIGHT_BOLD)
			wrapperFont.setBold();

		if (poiFont.getItalic())
			wrapperFont.setItalic();

		final short sizeInTwentiethOfPoint = poiFont.getFontHeight();
		final double sizeInPoints = sizeInTwentiethOfPoint / 20.0;
		if (!Util.almostEqual(sizeInPoints, 10.0))
			wrapperFont.setSize(sizeInPoints);

		final Map<Integer, HSSFColor> indexHash = HSSFColor.getIndexHash();
		final short fontColorIndex = poiFont.getColor();
		final HSSFColor poiFontColor = indexHash.get(Integer
				.valueOf(fontColorIndex));
		if (this.colorByHssfColor.containsKey(poiFontColor)) {
			final WrapperColor fontColor = this.colorByHssfColor
					.get(poiFontColor);
			if (fontColor != null)
				wrapperFont.setColor(fontColor);
		}

		final String fontName = poiFont.getFontName();
		if (!XlsConstants.DEFAULT_FONT_NAME.equals(fontName))
			wrapperFont.setFamily(fontName);

		wrapperCellStyle.setCellFont(wrapperFont);

		WrapperColor wrapperColor = null;
		final short backgroundColorIndex = cellStyle.getFillForegroundColor();
		final HSSFColor poiBackgroundColor = indexHash.get(Integer
				.valueOf(backgroundColorIndex));
		if (cellStyle.getFillPattern() == CellStyle.SOLID_FOREGROUND
				&& this.colorByHssfColor.containsKey(poiBackgroundColor))
			wrapperColor = this.colorByHssfColor.get(poiBackgroundColor);
		if (WrapperColor.DEFAULT_BACKGROUND.equals(wrapperColor))
			wrapperColor = null;

		wrapperCellStyle.setBackgroundColor(wrapperColor);

		Borders borders = new Borders();
		final double borderWidth = this.getBorderLineSize(cellStyle);
		if (borderWidth != WrapperCellStyle.DEFAULT)
			borders.setLineWidth(borderWidth);
		
		wrapperCellStyle.setBorders(borders);
		return wrapperCellStyle;
	}

	private double getBorderLineSize(final CellStyle cellStyle) {
		final short borderBottom = cellStyle.getBorderBottom();
		final short borderTop = cellStyle.getBorderTop();
		final short borderLeft = cellStyle.getBorderLeft();
		final short borderRight = cellStyle.getBorderRight();

		for (final short border : Arrays.asList(borderTop, borderLeft,
				borderRight)) {
			if (border != borderBottom)
				return WrapperCellStyle.DEFAULT;
		}

		final short borderBottomColor = cellStyle.getBottomBorderColor();
		final short borderTopColor = cellStyle.getTopBorderColor();
		final short borderLeftColor = cellStyle.getLeftBorderColor();
		final short borderRightColor = cellStyle.getRightBorderColor();

		for (final short borderColor : Arrays.asList(borderTopColor,
				borderLeftColor, borderRightColor)) {
			if (borderColor != borderBottomColor)
				return WrapperCellStyle.DEFAULT;
		}

		// TODO : check if black

		if (Util.almostEqual(CellStyle.BORDER_THIN, borderBottom))
			return WrapperCellStyle.THIN_LINE;
		else if (Util.almostEqual(CellStyle.BORDER_MEDIUM, borderBottom))
			return WrapperCellStyle.MEDIUM_LINE;
		else if (Util.almostEqual(CellStyle.BORDER_THICK, borderBottom))
			return WrapperCellStyle.THICK_LINE;
		else
			return WrapperCellStyle.DEFAULT;
	}

	private void setAllborders(final CellStyle cellStyle, final short border) {
		cellStyle.setBorderBottom(border);
		cellStyle.setBorderTop(border);
		cellStyle.setBorderLeft(border);
		cellStyle.setBorderRight(border);
	}

}
