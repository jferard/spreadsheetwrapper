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

import java.util.HashMap;
import java.util.Map;

import jxl.format.Colour;

import org.apache.poi.hssf.util.HSSFColor;

public class WrapperCellStyle {
	public static enum Color {
		AQUA(Colour.AQUA, new HSSFColor.AQUA()), AUTOMATIC(Colour.AUTOMATIC,
				new HSSFColor.AUTOMATIC()), BLACK(Colour.BLACK,
						new HSSFColor.BLACK()), BLUE(Colour.BLUE, new HSSFColor.BLUE());

		private String hex;
		private HSSFColor hssfColor;
		private Colour jxlColor;

		Color(final Colour jxlColor, final HSSFColor hssfColor) {
			this.jxlColor = jxlColor;
			this.hssfColor = hssfColor;
			final short[] triplet = this.hssfColor.getTriplet();
			final StringBuilder sb = new StringBuilder("#");
			for (final short l : triplet)
				sb.append(Integer.toHexString(l));
			this.hex = sb.toString();

			WrapperCellStyle.colorByHex.put(this.hex, this);
			WrapperCellStyle.colorByJxl.put(this.jxlColor, this);
			WrapperCellStyle.colorByHssf.put(this.hssfColor, this);
		}

		public HSSFColor getHssfColor() {
			return this.hssfColor;
		}

		public Colour getJxlColor() {
			return this.jxlColor;
		}

		public String toHex() {
			return this.hex;
		}
	}

	public static Map<String, Color> colorByHex = new HashMap<String, Color>();
	public static Map<org.apache.poi.ss.usermodel.Color, Color> colorByHssf = new HashMap<org.apache.poi.ss.usermodel.Color, Color>();

	public static Map<Colour, Color> colorByJxl = new HashMap<Colour, Color>();

	private Color backgoundColor;
	private WrapperFont cellFont;

	public WrapperCellStyle(final Color color, final WrapperFont cellFont) {
		super();
		this.backgoundColor = color;
		this.cellFont = cellFont;
	};

	public Color getBackgroundColor() {
		return this.backgoundColor;
	}

	public WrapperFont getCellFont() {
		return this.cellFont;
	}

	public void setBackgroundColor(final Color color) {
		this.backgoundColor = color;
	}

	public void setCellFont(final WrapperFont cellFont) {
		this.cellFont = cellFont;
	}
}
