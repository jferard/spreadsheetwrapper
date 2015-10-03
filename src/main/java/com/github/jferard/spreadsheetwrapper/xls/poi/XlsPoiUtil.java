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

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

public class XlsPoiUtil {
	public String getStyleString(final Workbook workbook,
			final CellStyle cellStyle) {
		final StringBuilder sb = new StringBuilder();
		final short fontIndex = cellStyle.getFontIndex();
		final Font font = workbook.getFontAt(fontIndex);
		if (font.getBoldweight() == Font.BOLDWEIGHT_BOLD)
			sb.append("font-weight:bold;");
		final Color color = cellStyle.getFillBackgroundColorColor();
		return "";
	}

	com.github.jferard.spreadsheetwrapper.WrapperCellStyle getCellStyle(
			final Workbook workbook, final CellStyle cellStyle) {
		final short fontIndex = cellStyle.getFontIndex();
		final Font font = workbook.getFontAt(fontIndex);
		com.github.jferard.spreadsheetwrapper.WrapperCellStyle c = null;
		com.github.jferard.spreadsheetwrapper.WrapperFont f = null;
		if (font.getBoldweight() == Font.BOLDWEIGHT_BOLD)
			f = new com.github.jferard.spreadsheetwrapper.WrapperFont(true, false,
					fontIndex, null);
		final Color color = cellStyle.getFillBackgroundColorColor();
		com.github.jferard.spreadsheetwrapper.WrapperCellStyle.Color col = null;
		if (color != null)
			col = com.github.jferard.spreadsheetwrapper.WrapperCellStyle.colorByHssf
			.get(color);

		c = new com.github.jferard.spreadsheetwrapper.WrapperCellStyle(col, f);
		return c;
	}
}
