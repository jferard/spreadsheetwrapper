package com.github.jferard.spreadsheetwrapper.xls.jxl;

import jxl.format.BoldStyle;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.format.Font;
import jxl.write.WritableFont;
import jxl.write.WriteException;

import com.github.jferard.spreadsheetwrapper.Util;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;
import com.github.jferard.spreadsheetwrapper.xls.XlsConstants;

class XlsJxlStyleFontHelper {
	private XlsJxlStyleColorHelper colorHelper;

	XlsJxlStyleFontHelper(XlsJxlStyleColorHelper colorHelper) {
		this.colorHelper = colorHelper;
	}

	public WritableFont toCellFont(WrapperCellStyle wrapperCellStyle)
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
			cellFont.setColour(this.colorHelper.toJxlColor(fontColor));
	
		return cellFont;
	}

	public WrapperFont toWrapperFont(final CellFormat cellFormat) {
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
				final WrapperColor wrapperFontColor = this.colorHelper
						.toWrapperColor(fontColor);
				if (wrapperFontColor != null)
					wrapperFont.setColor(wrapperFontColor);
			}
	
			final String name = font.getName();
			if (!XlsConstants.DEFAULT_FONT_NAME.equals(name))
				wrapperFont.setFamily(name);
		}
		return wrapperFont;
	}
}
