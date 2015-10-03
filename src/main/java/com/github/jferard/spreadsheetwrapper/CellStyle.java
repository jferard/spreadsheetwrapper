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
package com.github.jferard.spreadsheetwrapper;

import java.util.Map;

import jxl.format.Colour;

import org.apache.poi.hssf.util.HSSFColor;

public class CellStyle {
	public static Map<String, Color> colorByHex;
	public static Map<Colour, Color> colorByJxl;
	public static Map<org.apache.poi.ss.usermodel.Color, Color> colorByHssf;

	public static enum Color {
		AQUA(Colour.AQUA, new HSSFColor.AQUA()), AUTOMATIC(Colour.AUTOMATIC,
				new HSSFColor.AUTOMATIC()), BLACK(Colour.BLACK, new HSSFColor.BLACK()), BLUE(
						Colour.BLUE, new HSSFColor.BLUE());

		private String hex;
		private HSSFColor hssfColor;
		private Colour jxlColor;

		Color(final Colour jxlColor, final HSSFColor hssfColor) {
			this.jxlColor = jxlColor;
			this.hssfColor = hssfColor;
			final short[] triplet = this.hssfColor.getTriplet();
			final StringBuilder sb = new StringBuilder('#');
			for (final short l : triplet)
				sb.append(Integer.toHexString(l));
			this.hex = sb.toString();
			
			CellStyle.colorByHex.put(this.hex, this);
			CellStyle.colorByJxl.put(this.jxlColor, this);
			CellStyle.colorByHssf.put(this.hssfColor, this);
		}

		public String toHex() {
			return this.hex;
		}

		public HSSFColor getHssfColor() {
			return this.hssfColor;
		}

		public Colour getJxlColor() {
			return this.jxlColor;
		}
	}

	private Color backgoundColor;
	private Font cellFont;

	public CellStyle(final Color color, Font cellFont) {
		super();
		this.backgoundColor = color;
		this.cellFont = cellFont;
	};

	public Color getBackgroundColor() {
		return this.backgoundColor;
	}

	public void setBackgroundColor(final Color color) {
		this.backgoundColor = color;
	}

	public Font getCellFont() {
		return this.cellFont;
	}

	public void setCellFont(Font cellFont) {
		this.cellFont = cellFont;
	}
}
