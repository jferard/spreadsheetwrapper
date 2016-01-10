package com.github.jferard.spreadsheetwrapper.xls.poi;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.jferard.spreadsheetwrapper.Util;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;
import com.github.jferard.spreadsheetwrapper.xls.XlsConstants;

class XlsPoiStyleFontHelper {
	private XlsPoiStyleColorHelper colorHelper;

	XlsPoiStyleFontHelper(XlsPoiStyleColorHelper colorHelper) {
		this.colorHelper = colorHelper;
	}

	public Font toCellFont(final Workbook workbook,
			final WrapperCellStyle wrapperCellStyle) {
		final WrapperFont wrapperFont = wrapperCellStyle.getCellFont();
		if (wrapperFont == null)
			return null;
	
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
			final HSSFColor hssfColor = this.colorHelper.toHSSFColor(fontColor);
			final short index = hssfColor.getIndex();
			font.setColor(index);
		}
	
		if (fontFamily != null)
			font.setFontName(fontFamily);
	
		return font;
	}

	public WrapperFont toWrapperFont(Workbook workbook, CellStyle cellStyle) {
		final short fontIndex = cellStyle.getFontIndex();
		final Font poiFont = workbook.getFontAt(fontIndex);
		if (poiFont == null)
			return null;
	
		final WrapperFont wrapperFont = new WrapperFont();
		if (poiFont.getBoldweight() == Font.BOLDWEIGHT_BOLD)
			wrapperFont.setBold();
	
		if (poiFont.getItalic())
			wrapperFont.setItalic();
	
		final short sizeInTwentiethOfPoint = poiFont.getFontHeight();
		final double sizeInPoints = sizeInTwentiethOfPoint / 20.0;
		if (!Util.almostEqual(sizeInPoints, 10.0))
			wrapperFont.setSize(sizeInPoints);
	
		final short fontColorIndex = poiFont.getColor();
		final HSSFColor poiFontColor = this.colorHelper.getHSSFColor(fontColorIndex);
		final WrapperColor fontColor = this.colorHelper.toWrapperColor(poiFontColor);
		if (fontColor != null)
			wrapperFont.setColor(fontColor);
	
		final String fontName = poiFont.getFontName();
		if (!XlsConstants.DEFAULT_FONT_NAME.equals(fontName))
			wrapperFont.setFamily(fontName);
	
		return wrapperFont;
	}

}
