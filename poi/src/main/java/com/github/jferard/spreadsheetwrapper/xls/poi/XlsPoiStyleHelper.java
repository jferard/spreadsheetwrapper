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

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.jferard.spreadsheetwrapper.style.Borders;
import com.github.jferard.spreadsheetwrapper.style.CellStyleAccessor;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * An utility class for style handling
 *
 */
class XlsPoiStyleHelper {
	/** name or index -> cellStyle */
	private final CellStyleAccessor<CellStyle> cellStyleAccessor;
	private XlsPoiStyleColorHelper colorHelper;
	private XlsPoiStyleFontHelper fontHelper;
	private XlsPoiStyleBorderHelper borderHelper;

	// private HSSFColor hssfColor;

	/**
	 * @param cellStyleAccessor
	 */
	public XlsPoiStyleHelper(
			final CellStyleAccessor<CellStyle> cellStyleAccessor,
			XlsPoiStyleColorHelper colorHelper,
			XlsPoiStyleFontHelper fontHelper,
			XlsPoiStyleBorderHelper borderHelper) {
		this.cellStyleAccessor = cellStyleAccessor;
		this.colorHelper = colorHelper;
		this.fontHelper = fontHelper;
		this.borderHelper = borderHelper;
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
		final Font font = this.fontHelper
				.toCellFont(workbook, wrapperCellStyle);
		cellStyle.setFont(font);

		final WrapperColor backgroundColor = wrapperCellStyle
				.getBackgroundColor();

		if (backgroundColor != null) {
			final HSSFColor hssfColor = this.colorHelper.toHSSFColor(backgroundColor);

			final short index = hssfColor.getIndex();
			cellStyle.setFillForegroundColor(index);
			cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		}

		this.borderHelper.setCellBorders(wrapperCellStyle, cellStyle);
		return cellStyle;
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

		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle();
		WrapperFont wrapperFont = this.fontHelper.toWrapperFont(workbook, cellStyle);
		if (wrapperFont != null)
			wrapperCellStyle.setCellFont(wrapperFont);

		WrapperColor wrapperColor = this.colorHelper.toWrapperColor(cellStyle);
		if (wrapperColor != null)
			wrapperCellStyle.setBackgroundColor(wrapperColor);

		Borders borders = this.borderHelper.toBorders(cellStyle);
		if (borders != null)
			wrapperCellStyle.setBorders(borders);

		return wrapperCellStyle;
	}

}
