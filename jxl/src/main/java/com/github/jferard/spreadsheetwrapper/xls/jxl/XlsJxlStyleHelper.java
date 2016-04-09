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
package com.github.jferard.spreadsheetwrapper.xls.jxl;

import java.util.HashMap;

import jxl.Cell;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WriteException;

import com.github.jferard.spreadsheetwrapper.StyleHelper;
import com.github.jferard.spreadsheetwrapper.style.Borders;
import com.github.jferard.spreadsheetwrapper.style.CellStyleAccessor;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

class XlsJxlStyleHelper implements StyleHelper<CellFormat, WritableCell> {
	/** name or index -> cellStyle */
	private final CellStyleAccessor<WritableCellFormat> cellStyleAccessor;
	private XlsJxlStyleFontHelper fontHelper;
	private XlsJxlStyleBorderHelper borderHelper;
	private XlsJxlStyleColorHelper colorHelper;

	XlsJxlStyleHelper(
			final CellStyleAccessor<WritableCellFormat> cellStyleAccessor,
			XlsJxlStyleColorHelper colorHelper,
			XlsJxlStyleFontHelper fontHelper,
			XlsJxlStyleBorderHelper borderHelper) {
		this.cellStyleAccessor = cellStyleAccessor;
		this.colorHelper = colorHelper;
		this.fontHelper = fontHelper;
		this.borderHelper = borderHelper;
	}

	/**
	 * @param styleName
	 *            the name of the style
	 * @return the internal style
	 */
	public/*@Nullable*/WritableCellFormat getCellFormat(
			final String styleName) {
		final WritableCellFormat cellFormat = this.cellStyleAccessor
				.getCellStyle(styleName);
		return cellFormat;
	}

	// /**
	// * @param cellFormat
	// * the internal style
	// * @return the style name, ssw<index> if none
	// */
	// public String getStyleName(final WritableCellFormat cellFormat) {
	// final String name = this.cellStyleAccessor.getName(cellFormat);
	// if (name == null)
	// return "ssw" + cellFormat.getFormatRecord();
	// else
	// return name;
	// }

	/**
	 * Create or update a cell style
	 *
	 * @param styleName
	 *            the name of the style
	 * @param cellFormat
	 *            the internal style
	 */
	public void putCellStyle(final String styleName,
			final WritableCellFormat cellFormat) {
		this.cellStyleAccessor.putCellStyle(styleName, cellFormat);
	}

	/**
	 * @param wrapperCellStyle
	 *            the wrapper cell style
	 * @return the internal cell style
	 */
	public WritableCellFormat toCellFormat(
			final WrapperCellStyle wrapperCellStyle) {
		final WritableCellFormat cellFormat = new WritableCellFormat();
		try {
			WritableFont /*@Nullable*/ cellFont = this.fontHelper
					.toCellFont(wrapperCellStyle);
			if (cellFont != null)
				cellFormat.setFont(cellFont);

			final WrapperColor backgroundColor = wrapperCellStyle
					.getBackgroundColor();
			if (backgroundColor != null) {
				final Colour jxlColor = this.colorHelper
						.toJxlColor(backgroundColor);
				if (jxlColor != null)
					cellFormat.setBackground(jxlColor);
			}

			final/*@Nullable*/Borders borders = wrapperCellStyle.getBorders();
			if (borders != null)
				this.borderHelper.setBorders(cellFormat, borders);
		} catch (final WriteException e) {
			throw new IllegalStateException(e);
		}
		return cellFormat;
	}

	/**
	 * @param cellFormat
	 *            the internal cell style
	 * @return the wrapper cell style
	 */
	@Override
	public WrapperCellStyle toWrapperCellStyle(final CellFormat cellFormat) {
		if (cellFormat == null)
			return WrapperCellStyle.EMPTY;
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle();

		WrapperColor backgroundColor = this.colorHelper
				.toWrapperColor(cellFormat.getBackgroundColour());
		if (WrapperColor.DEFAULT_BACKGROUND.equals(backgroundColor))
			backgroundColor = null;
		if (backgroundColor != null)
			wrapperCellStyle.setBackgroundColor(backgroundColor);

		final Borders borders = this.borderHelper.getBorders(cellFormat);
		if (borders != null)
			wrapperCellStyle.setBorders(borders);

		final WrapperFont wrapperFont = this.fontHelper
				.toWrapperFont(cellFormat);
		if (wrapperFont != null)
			wrapperCellStyle.setCellFont(wrapperFont);

		return wrapperCellStyle;
	}

	@Override
	public WrapperCellStyle getWrapperCellStyle(WritableCell element) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void setWrapperCellStyle(WritableCell cell,
			WrapperCellStyle wrapperCellStyle) {
		cell.setCellFormat(this.toCellFormat(wrapperCellStyle));
	}

	public WrapperCellStyle getCellStyle(final String styleName) {
		final WritableCellFormat cellFormat = this.getCellFormat(styleName);
		if (cellFormat == null)
			return WrapperCellStyle.EMPTY;

		return this.toWrapperCellStyle(cellFormat);
	}

	public boolean setStyle(final String styleName,
			final WrapperCellStyle wrapperCellStyle) {
		final WritableCellFormat cellFormat = this
				.toCellFormat(wrapperCellStyle);
		this.putCellStyle(styleName, cellFormat);
		return true;
	}

	public WrapperCellStyle getStyle(Cell cell) {
		if (cell == null)
			return null;

		final CellFormat cellFormat = cell.getCellFormat();
		return this.toWrapperCellStyle(cellFormat);
	}

	public boolean setStyle(WritableCell cell,
			WrapperCellStyle wrapperCellStyle) {
		if (cell == null)
			return false;

		this.setWrapperCellStyle(cell, wrapperCellStyle);
		return true;
	}

	public boolean setStyleName(WritableCell cell, String styleName) {
		final WritableCellFormat cellFormat = this.getCellFormat(styleName);
		if (cellFormat == null)
			return false;
		else {
			cell.setCellFormat(cellFormat);
			return true;
		}
	}

}
