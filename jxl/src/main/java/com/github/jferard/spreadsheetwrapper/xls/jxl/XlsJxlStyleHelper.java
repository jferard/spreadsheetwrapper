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

import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.format.BoldStyle;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.format.Font;
import jxl.format.RGB;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WriteException;

import com.github.jferard.spreadsheetwrapper.StyleHelper;
import com.github.jferard.spreadsheetwrapper.Util;
import com.github.jferard.spreadsheetwrapper.style.Borders;
import com.github.jferard.spreadsheetwrapper.style.CellStyleAccessor;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;
import com.github.jferard.spreadsheetwrapper.xls.XlsConstants;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class XlsJxlStyleHelper implements StyleHelper<CellFormat, WritableCell> {
	private static boolean colorEquals(final Colour color1, final Colour color2) {
		final RGB rgb1 = color1.getDefaultRGB();
		final RGB rgb2 = color2.getDefaultRGB();
		return rgb1.getRed() == rgb2.getRed()
				&& rgb1.getGreen() == rgb2.getGreen()
				&& rgb1.getBlue() == rgb2.getBlue();
	}

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
			WritableFont /*@Nullable*/cellFont = this
					.toCellFont(wrapperCellStyle);
			if (cellFont != null)
				cellFormat.setFont(cellFont);

			final WrapperColor backgroundColor = wrapperCellStyle
					.getBackgroundColor();
			if (backgroundColor != null) {
				final Colour jxlColor = this.toJxlColor(backgroundColor);
				if (jxlColor != null)
					cellFormat.setBackground(jxlColor);
			}

			final/*@Nullable*/Borders borders = wrapperCellStyle.getBorders();
			if (borders != null)
				this.setBorders(cellFormat, borders);
		} catch (final WriteException e) {
			throw new IllegalStateException(e);
		}
		return cellFormat;
	}

	private WritableFont toCellFont(WrapperCellStyle wrapperCellStyle)
			throws WriteException {
		WritableFont cellFont;

		final WrapperFont wrapperFont = wrapperCellStyle.getCellFont();
		if (wrapperFont == null)
			return null;

		final String family = wrapperFont.getFamily();
		if (family == null)
			cellFont = new WritableFont(WritableFont.ARIAL);
		else if (WrapperFont.COURIER_NAME.equals(family))
			cellFont = new WritableFont(WritableFont.COURIER);
		else if (WrapperFont.TAHOMA_NAME.equals(family))
			cellFont = new WritableFont(WritableFont.TAHOMA);
		else if (WrapperFont.TIMES_NAME.equals(family))
			cellFont = new WritableFont(WritableFont.TIMES);
		else
			cellFont = new WritableFont(WritableFont.ARIAL);

		final int bold = wrapperFont.getBold();
		if (bold == WrapperCellStyle.YES)
			cellFont.setBoldStyle(WritableFont.BOLD);
		else if (bold == WrapperCellStyle.NO)
			cellFont.setBoldStyle(WritableFont.NO_BOLD);

		final int italic = wrapperFont.getItalic();
		if (italic == WrapperCellStyle.YES)
			cellFont.setItalic(true);
		else if (italic == WrapperCellStyle.NO)
			cellFont.setItalic(false);

		final double size = wrapperFont.getSize();
		if (!Util.almostEqual(size, WrapperCellStyle.DEFAULT))
			cellFont.setPointSize((int) size);

		final WrapperColor fontColor = wrapperFont.getColor();
		if (fontColor != null)
			cellFont.setColour(this.toJxlColor(fontColor));

		return cellFont;
	}

	private void setBorders(final WritableCellFormat cellFormat,
			final Borders borders) throws WriteException {
		final double borderLineWidth = borders.getLineWidth();
		WrapperColor lineColor = borders.getLineColor();

		Colour color = lineColor == null ? Colour.BLACK : this
				.toJxlColor(lineColor);
		BorderLineStyle style;
		if (borderLineWidth != WrapperCellStyle.DEFAULT) {
			if (Util.almostEqual(borderLineWidth, WrapperCellStyle.THIN_LINE))
				style = BorderLineStyle.THIN;
			else if (Util.almostEqual(borderLineWidth,
					WrapperCellStyle.MEDIUM_LINE))
				style = BorderLineStyle.MEDIUM;
			else if (Util.almostEqual(borderLineWidth,
					WrapperCellStyle.THICK_LINE))
				style = BorderLineStyle.THICK;
			else
				throw new UnsupportedOperationException();

			cellFormat.setBorder(Border.ALL, style, color);
		}
	}

	/**
	 * @param wrapperColor
	 *            the wrapper color
	 * @return the internal color
	 */
	private/*@Nullable*/Colour toJxlColor(final WrapperColor wrapperColor) {
		return this.jxlColorByWrapperColor.get(wrapperColor);
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

		WrapperColor backgroundColor = this.toWrapperColor(cellFormat
				.getBackgroundColour());
		if (WrapperColor.DEFAULT_BACKGROUND.equals(backgroundColor))
			backgroundColor = null;
		if (backgroundColor != null)
			wrapperCellStyle.setBackgroundColor(backgroundColor);

		final Borders borders = this.getBorders(cellFormat);
		wrapperCellStyle.setBorders(borders);

		final WrapperFont wrapperFont = new WrapperFont();
		final Font font = cellFormat.getFont();

		if (font != null) {
			final int boldWeight = font.getBoldWeight();
			if (boldWeight == BoldStyle.BOLD.getValue())
				wrapperFont.setBold();

			if (font.isItalic())
				wrapperFont.setItalic();

			final int pointSize = font.getPointSize();
			if (pointSize != WritableFont.DEFAULT_POINT_SIZE)
				wrapperFont.setSize(pointSize);

			final Colour fontColor = font.getColour();
			if (fontColor != Colour.BLACK && fontColor != Colour.AUTOMATIC) {
				final WrapperColor wrapperFontColor = this
						.toWrapperColor(fontColor);
				if (wrapperFontColor != null)
					wrapperFont.setColor(wrapperFontColor);
			}

			final String name = font.getName();
			if (!XlsConstants.DEFAULT_FONT_NAME.equals(name))
				wrapperFont.setFamily(name);
		}

		return wrapperCellStyle.setCellFont(wrapperFont);
	}

	/**
	 * @param jxlColor
	 *            internal color
	 * @return wrapper color
	 */
	public/*@Nullable*/WrapperColor toWrapperColor(final Colour jxlColor) {
		return this.wrapperColorByJxlColor.get(jxlColor);
	}

	private Borders getBorders(final CellFormat cellFormat) {
		Borders borders = new Borders();
		Double borderLineWidth = this.getBorderLineWidth(cellFormat);
		if (borderLineWidth != null)
			borders.setLineWidth(borderLineWidth);

		WrapperColor borderColor = this.getBorderColor(cellFormat);
		if (borderColor != null)
			borders.setLineColor(borderColor);

		return borders;
	}

	private Double getBorderLineWidth(CellFormat cellFormat) {
		final List<Border> list = Arrays.asList(Border.RIGHT, Border.TOP,
				Border.BOTTOM);
		final BorderLineStyle borderStyle = cellFormat
				.getBorderLine(Border.LEFT);
		for (final Border border : list) {
			if (!borderStyle.equals(cellFormat.getBorderLine(border))) {
				return null;
			}
		}
		final Double borderLineWidth;
		if (BorderLineStyle.THIN.equals(borderStyle))
			borderLineWidth = WrapperCellStyle.THIN_LINE;
		else if (BorderLineStyle.MEDIUM.equals(borderStyle))
			borderLineWidth = WrapperCellStyle.MEDIUM_LINE;
		else if (BorderLineStyle.THICK.equals(borderStyle))
			borderLineWidth = WrapperCellStyle.THICK_LINE;
		else
			borderLineWidth = null;
		return borderLineWidth;
	}

	private WrapperColor getBorderColor(CellFormat cellFormat) {
		final List<Border> list = Arrays.asList(Border.RIGHT, Border.TOP,
				Border.BOTTOM);
		final Colour borderColor = cellFormat.getBorderColour(Border.LEFT);
		for (final Border border : list) {
			if (!XlsJxlStyleHelper.colorEquals(borderColor,
					cellFormat.getBorderColour(border))) {
				return null;
			}
		}

		if (XlsJxlStyleHelper.colorEquals(borderColor, Colour.BLACK)
				|| XlsJxlStyleHelper.colorEquals(borderColor,
						Colour.PALETTE_BLACK)
				|| XlsJxlStyleHelper.colorEquals(borderColor, Colour.AUTOMATIC))
			return null;

		return this.toWrapperColor(borderColor);
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
}
