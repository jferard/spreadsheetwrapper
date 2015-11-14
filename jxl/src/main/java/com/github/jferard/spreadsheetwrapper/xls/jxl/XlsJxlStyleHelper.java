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

import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.WrapperFont;
import com.github.jferard.spreadsheetwrapper.impl.CellStyleAccessor;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class XlsJxlStyleHelper {
	private static final int BOLDWEIGHT_BOLD = 400;
	private final CellStyleAccessor<WritableCellFormat> cellStyleAccessor;
	private final Map<WrapperColor, Colour> jxlColorByWrapperColor;
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
				if (jxlColor != null)
					this.jxlColorByWrapperColor.put(color, jxlColor);
				this.wrapperColorByJxlColor.put(jxlColor, color);
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

	public/*@Nullable*/Colour getJxlColor(final WrapperColor color) {
		return this.jxlColorByWrapperColor.get(color);
	}

	public String getStyleName(final WritableCellFormat cellFormat) {
		final String name = this.cellStyleAccessor.getName(cellFormat);
		if (name == null)
			return "ssw" + cellFormat.getFormatRecord();
		else
			return name;
	}

	public WrapperCellStyle getWrapperCellStyle(
			final WritableCellFormat cellFormat) {
		final WrapperColor backgroundColor = this
				.getWrapperColor(cellFormat.getBackgroundColour());
		final WrapperFont wrapperFont = new WrapperFont();
		final Font font = cellFormat.getFont();
		final Colour fontColor = font.getColour();
		if (fontColor != Colour.BLACK && fontColor != Colour.AUTOMATIC)
			wrapperFont
					.setColor(this.getWrapperColor(fontColor));
		if (font.getBoldWeight() == BoldStyle.BOLD.getValue())
			wrapperFont.setBold();

		return new WrapperCellStyle(backgroundColor, wrapperFont);
	}

	public WrapperCellStyle getWrapperCellStyle(
			final CellFormat cellFormat) {
		final WrapperColor backgroundColor = this
				.getWrapperColor(cellFormat.getBackgroundColour());
		final WrapperFont wrapperFont = new WrapperFont();
		final Font font = cellFormat.getFont();
		final Colour fontColor = font.getColour();
		if (fontColor != Colour.BLACK && fontColor != Colour.AUTOMATIC)
			wrapperFont
					.setColor(this.getWrapperColor(fontColor));
		if (font.getBoldWeight() == BoldStyle.BOLD.getValue())
			wrapperFont.setBold();

		return new WrapperCellStyle(backgroundColor, wrapperFont);
	}
	
	public/*@Nullable*/WrapperColor getWrapperColor(final Colour color) {
		return this.wrapperColorByJxlColor.get(color);
	}

	public void putCellStyle(final String styleName,
			final WritableCellFormat cellFormat) {
		this.cellStyleAccessor.putCellStyle(styleName, cellFormat);
	}

	WritableCellFormat getCellFormat(final String styleName) {
		final WritableCellFormat cellFormat = this.cellStyleAccessor
				.getCellStyle(styleName);
		// WorkbookParser parser = (WorkbookParser) workbook;
		// if (cellFormat == null) {
		// if (styleName.startsWith("ssw")) {
		// try {
		// int idx = Integer.valueOf(styleName.substring(3));
		// FormattingRecords formattingRecords = parser
		// .getFormattingRecords();
		// cellFormat = (WritableCellFormat) formattingRecords
		// .getXFRecord(idx);
		// } catch (NumberFormatException e) {
		// // do nothing
		// }
		// }
		// }
		return cellFormat;
	}

	WritableCellFormat getCellFormat(final WrapperCellStyle wrapperCellStyle) {
		final WritableFont cellFont = new WritableFont(WritableFont.ARIAL);
		final WritableCellFormat cellFormat = new WritableCellFormat(cellFont);
		try {
			final WrapperFont wrapperFont = wrapperCellStyle.getCellFont();
			if (wrapperFont != null) {
				if (wrapperFont.getBold() == WrapperCellStyle.YES)
					cellFont.setBoldStyle(WritableFont.BOLD);
			}
			final WrapperColor backgroundColor = wrapperCellStyle
					.getBackgroundColor();
			if (backgroundColor != null) {
				final Colour jxlColor = this.getJxlColor(backgroundColor);
				if (jxlColor != null)
					cellFormat.setBackground(jxlColor);
			}
		} catch (final WriteException e) {
			// do nothing
		}
		return cellFormat;
	}

}
