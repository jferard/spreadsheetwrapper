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
import java.util.Map;

import jxl.format.BoldStyle;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.format.Font;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WriteException;

import com.github.jferard.spreadsheetwrapper.CellStyleAccessor;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class XlsJxlStyleHelper {
	/** name or index -> cellStyle */
	private final CellStyleAccessor<WritableCellFormat> cellStyleAccessor;
	/** wrapper -> internal */
	private final Map<WrapperColor, Colour> jxlColorByWrapperColor;
	/** internal -> wrapper */
	private final Map<Colour, WrapperColor> wrapperColorByJxlColor;

	XlsJxlStyleHelper(
			final CellStyleAccessor<WritableCellFormat> cellStyleAccessor) {
		this.cellStyleAccessor = cellStyleAccessor;
		final WrapperColor[] colors = WrapperColor.values();
		this.jxlColorByWrapperColor = new HashMap<WrapperColor, Colour>(
				colors.length);
		this.wrapperColorByJxlColor = new HashMap<Colour, WrapperColor>(
				colors.length);
		for (final WrapperColor color : colors) {
			Colour jxlColor;
			try {
				final Class<?> jxlClazz = Class.forName("jxl.format.Colour");
				jxlColor = (Colour) jxlClazz.getDeclaredField(
						color.getSimpleName()).get(null);
				if (jxlColor != null) {
					this.jxlColorByWrapperColor.put(color, jxlColor);
					this.wrapperColorByJxlColor.put(jxlColor, color);
				}
			} catch (final ClassNotFoundException e) {
				jxlColor = null;
			} catch (final IllegalArgumentException e) {
				jxlColor = null;
			} catch (final IllegalAccessException e) {
				jxlColor = null;
			} catch (final NoSuchFieldException e) {
				jxlColor = null;
			} catch (final SecurityException e) {
				jxlColor = null;
			}
		}
	}

	/**
	 * @param styleName
	 *            the name of the style
	 * @return the internal style
	 */
	public/*@Nullable*/WritableCellFormat getCellFormat(final String styleName) {
		final WritableCellFormat cellFormat = this.cellStyleAccessor
				.getCellStyle(styleName);
		return cellFormat;
	}

	/**
	 * @param cellFormat
	 *            the internal style
	 * @return the style name, ssw<index> if none
	 */
	public String getStyleName(final WritableCellFormat cellFormat) {
		final String name = this.cellStyleAccessor.getName(cellFormat);
		if (name == null)
			return "ssw" + cellFormat.getFormatRecord();
		else
			return name;
	}

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
		final WritableFont cellFont = new WritableFont(WritableFont.ARIAL);
		final WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
		try {
			final WrapperFont wrapperFont = wrapperCellStyle.getCellFont();
			if (wrapperFont != null) {
					if (wrapperFont.getBold() == WrapperCellStyle.YES)
						cellFont.setBoldStyle(WritableFont.BOLD);
					if (wrapperFont.getItalic() == WrapperCellStyle.YES)
						cellFont.setItalic(true);
					final WrapperColor fontColor = wrapperFont.getColor();
					if (fontColor != null)
						cellFont.setColour(this.toJxlColor(fontColor));
			}

			final WrapperColor backgroundColor = wrapperCellStyle
					.getBackgroundColor();
			if (backgroundColor != null) {
				final Colour jxlColor = this.toJxlColor(backgroundColor);
				if (jxlColor != null)
					cellFormat.setBackground(jxlColor);
			}
		} catch (final WriteException e) { // NOPMD by Julien on 24/11/15 19:48
			// do nothing
		}
		return cellFormat;
	}

	/**
	 * @param wrapperColor
	 *            the wrapper color
	 * @return the internal color
	 */
	public/*@Nullable*/Colour toJxlColor(final WrapperColor wrapperColor) {
		return this.jxlColorByWrapperColor.get(wrapperColor);
	}

	/**
	 * @param cellFormat
	 *            the internal cell style
	 * @return the wrapper cell style
	 */
	public WrapperCellStyle toWrapperCellStyle(final CellFormat cellFormat) {
		if (cellFormat == null)
			return WrapperCellStyle.EMPTY;

		WrapperColor backgroundColor = this.toWrapperColor(cellFormat
				.getBackgroundColour());
		if (WrapperColor.DEFAULT_BACKGROUND.equals(backgroundColor))
			backgroundColor = null;

		final WrapperFont wrapperFont = new WrapperFont();
		final Font font = cellFormat.getFont();
		final Colour fontColor = font.getColour();
		if (fontColor != Colour.BLACK && fontColor != Colour.AUTOMATIC) {
			final WrapperColor wrapperFontColor = this
					.toWrapperColor(fontColor);
			if (wrapperFontColor != null)
				wrapperFont.setColor(wrapperFontColor);
		}
		if (font.getBoldWeight() == BoldStyle.BOLD.getValue())
			wrapperFont.setBold();

		return new WrapperCellStyle(backgroundColor, wrapperFont);
	}

	/**
	 * @param cellFormat
	 *            the internal cell style
	 * @return the wrapper cell style
	 */
	public WrapperCellStyle toWrapperCellStyle(
			final WritableCellFormat cellFormat) {
		if (cellFormat == null)
			return WrapperCellStyle.EMPTY;

		final WrapperColor backgroundColor = this.toWrapperColor(cellFormat
				.getBackgroundColour());
		final WrapperFont wrapperFont = new WrapperFont();
		final Font font = cellFormat.getFont();
		final Colour fontColor = font.getColour();
		if (fontColor != Colour.BLACK && fontColor != Colour.AUTOMATIC) {
			final WrapperColor wrapperFontColor = this
					.toWrapperColor(fontColor);
			if (wrapperFontColor != null)
				wrapperFont.setColor(wrapperFontColor);
		}
		if (font.getBoldWeight() == BoldStyle.BOLD.getValue())
			wrapperFont.setBold();

		return new WrapperCellStyle(backgroundColor, wrapperFont);
	}

	/**
	 * @param jxlColor
	 *            internal color
	 * @return wrapper color
	 */
	public/*@Nullable*/WrapperColor toWrapperColor(final Colour jxlColor) {
		return this.wrapperColorByJxlColor.get(jxlColor);
	}
}
