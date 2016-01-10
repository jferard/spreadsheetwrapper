package com.github.jferard.spreadsheetwrapper.xls.jxl;

import java.util.Arrays;
import java.util.List;

import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.write.WritableCellFormat;
import jxl.write.WriteException;

import com.github.jferard.spreadsheetwrapper.Util;
import com.github.jferard.spreadsheetwrapper.style.Borders;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;

class XlsJxlStyleBorderHelper {
	private XlsJxlStyleColorHelper colorHelper;

	XlsJxlStyleBorderHelper(XlsJxlStyleColorHelper colorHelper) {
		this.colorHelper = colorHelper;
	}

	private WrapperColor getBorderColor(CellFormat cellFormat) {
		final List<Border> list = Arrays.asList(Border.RIGHT, Border.TOP,
				Border.BOTTOM);
		final Colour borderColor = cellFormat.getBorderColour(Border.LEFT);
		for (final Border border : list) {
			if (!XlsJxlStyleColorHelper.colorEquals(borderColor,
					cellFormat.getBorderColour(border))) {
				return null;
			}
		}
	
		if (XlsJxlStyleColorHelper.colorEquals(borderColor, Colour.BLACK)
				|| XlsJxlStyleColorHelper.colorEquals(borderColor,
						Colour.PALETTE_BLACK)
				|| XlsJxlStyleColorHelper.colorEquals(borderColor, Colour.AUTOMATIC))
			return null;
	
		return this.colorHelper.toWrapperColor(borderColor);
	}

	private static Double getBorderLineWidth(CellFormat cellFormat) {
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

	public void setBorders(final WritableCellFormat cellFormat,
			final Borders borders) throws WriteException {
		final double borderLineWidth = borders.getLineWidth();
		WrapperColor lineColor = borders.getLineColor();
	
		Colour color = lineColor == null ? Colour.BLACK : this.colorHelper
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

	public Borders getBorders(CellFormat cellFormat) {
		Borders borders = new Borders();
		WrapperColor borderColor = this.getBorderColor(cellFormat);
		if (borderColor != null)
			borders.setLineColor(borderColor);
		
		Double borderLineWidth = XlsJxlStyleBorderHelper.getBorderLineWidth(cellFormat);
		if (borderLineWidth != null && borderLineWidth > 0.0)
			borders.setLineWidth(borderLineWidth);
		
		return borders;
	} 

}
