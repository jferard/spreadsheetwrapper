package com.github.jferard.spreadsheetwrapper.xls.poi;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

public class XlsPoiUtil {
	public String getStyleString(Workbook workbook, CellStyle cellStyle) {
		StringBuilder sb = new StringBuilder();
		short fontIndex = cellStyle.getFontIndex();
		Font font = workbook.getFontAt(fontIndex);
		if (font.getBoldweight() == Font.BOLDWEIGHT_BOLD)
			sb.append("font-weight:bold;");
		Color color = cellStyle.getFillBackgroundColorColor();
		return "";
	}
}
