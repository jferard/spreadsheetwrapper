package com.github.jferard.spreadsheetwrapper.xls.poi;

import java.util.Arrays;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;

import com.github.jferard.spreadsheetwrapper.Util;
import com.github.jferard.spreadsheetwrapper.style.Borders;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;

class XlsPoiStyleBorderHelper {
	
	private XlsPoiStyleColorHelper colorHelper;

	XlsPoiStyleBorderHelper(XlsPoiStyleColorHelper colorHelper) {
		this.colorHelper = colorHelper;
	}

	private WrapperColor getBorderLineColor(final CellStyle cellStyle) {
		final short borderBottomColor = cellStyle.getBottomBorderColor();
		final short borderTopColor = cellStyle.getTopBorderColor();
		final short borderLeftColor = cellStyle.getLeftBorderColor();
		final short borderRightColor = cellStyle.getRightBorderColor();
	
		for (final short borderColor : Arrays.asList(borderTopColor,
				borderLeftColor, borderRightColor)) {
			if (borderColor != borderBottomColor)
				return null;
		}
	
		HSSFColor hssfColor = this.colorHelper.getHSSFColor(borderBottomColor);
		if (hssfColor == null)
			return null;
		
		if (hssfColor.getHexString().equals("0:0:0"))
			return null;
	
		return this.colorHelper.toWrapperColor(hssfColor);
	}

	private static double getBorderLineSize(final CellStyle cellStyle) {
		final short borderBottom = cellStyle.getBorderBottom();
		final short borderTop = cellStyle.getBorderTop();
		final short borderLeft = cellStyle.getBorderLeft();
		final short borderRight = cellStyle.getBorderRight();
	
		for (final short border : Arrays.asList(borderTop, borderLeft,
				borderRight)) {
			if (border != borderBottom)
				return WrapperCellStyle.DEFAULT;
		}
	
		final double lineWidth;
		if (borderBottom == CellStyle.BORDER_THIN)
			lineWidth = WrapperCellStyle.THIN_LINE;
		else if (borderBottom == CellStyle.BORDER_MEDIUM)
			lineWidth = WrapperCellStyle.MEDIUM_LINE;
		else if (borderBottom == CellStyle.BORDER_THICK)
			lineWidth = WrapperCellStyle.THICK_LINE;
		else
			lineWidth = WrapperCellStyle.DEFAULT;
	
		return lineWidth;
	}

	private static void setAllBordersColor(final CellStyle cellStyle, final short color) {
		cellStyle.setBottomBorderColor(color);
		cellStyle.setTopBorderColor(color);
		cellStyle.setLeftBorderColor(color);
		cellStyle.setRightBorderColor(color);
	}

	private static void setAllBordersWidth(final CellStyle cellStyle,
			final short border) {
		cellStyle.setBorderBottom(border);
		cellStyle.setBorderTop(border);
		cellStyle.setBorderLeft(border);
		cellStyle.setBorderRight(border);
	}

	public void setCellBorders(final WrapperCellStyle wrapperCellStyle,
			final CellStyle cellStyle) {
		final/*@Nullable*/Borders borders = wrapperCellStyle.getBorders();
		if (borders == null)
			return;
	
		final double borderWidth = borders.getLineWidth();
		if (borderWidth != WrapperCellStyle.DEFAULT) {
			short lineWidth;
			if (Util.almostEqual(borderWidth, WrapperCellStyle.THIN_LINE))
				lineWidth = CellStyle.BORDER_THIN;
			else if (Util
					.almostEqual(borderWidth, WrapperCellStyle.MEDIUM_LINE))
				lineWidth = CellStyle.BORDER_MEDIUM;
			else if (Util.almostEqual(borderWidth, WrapperCellStyle.THICK_LINE))
				lineWidth = CellStyle.BORDER_THICK;
			else
				throw new UnsupportedOperationException();
	
			XlsPoiStyleBorderHelper.setAllBordersWidth(cellStyle, lineWidth);
		}
	
		final WrapperColor color = borders.getLineColor();
		if (color != null) {
			HSSFColor hssfColor = this.colorHelper.toHSSFColor(color);
			XlsPoiStyleBorderHelper.setAllBordersColor(cellStyle, hssfColor.getIndex());
		}
	}

	public Borders toBorders(CellStyle cellStyle) {
		Borders borders = new Borders();
		final double borderWidth = XlsPoiStyleBorderHelper.getBorderLineSize(cellStyle);
		if (borderWidth != WrapperCellStyle.DEFAULT)
			borders.setLineWidth(borderWidth);
	
		final WrapperColor color = this.getBorderLineColor(cellStyle);
		if (color != null)
			borders.setLineColor(color);
		return borders;
	
	}

}
