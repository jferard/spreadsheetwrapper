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

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class WrapperCellStyle {
	public static final int DEFAULT = -1;
	public static final WrapperCellStyle EMPTY = new WrapperCellStyle(null,
			new WrapperFont());
	public static final int NO = 0;
	public static final int YES = 1;
	private/*@Nullable*/WrapperColor backgoundColor;
	private/*@Nullable*/WrapperFont cellFont;

	/**
	 * @param backgroundColor
	 *            the background color of the cell
	 * @param cellFont
	 *            the font
	 */
	public WrapperCellStyle(final/*@Nullable*/WrapperColor backgroundColor,
			final/*@Nullable*/WrapperFont cellFont) {
		super();
		this.backgoundColor = backgroundColor;
		this.cellFont = cellFont;
	}

	@Override
	public boolean equals(final /*@Nullable*/ Object obj) {
		if (this == obj)
			return true;
		if (!(obj instanceof WrapperCellStyle))
			return false;

		final WrapperCellStyle other = (WrapperCellStyle) obj;
		return this.backgoundColor == other.backgoundColor
				&& Util.equal(this.cellFont, other.cellFont);
	}

	public/*@Nullable*/WrapperColor getBackgroundColor() {
		return this.backgoundColor;
	}

	public/*@Nullable*/WrapperFont getCellFont() {
		return this.cellFont;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		return prime * (prime + Util.hash(this.backgoundColor))
				+ Util.hash(this.cellFont);
	}

	public void setBackgroundColor(final WrapperColor wrapperColor) {
		this.backgoundColor = wrapperColor;
	}

	public void setCellFont(final WrapperFont cellFont) {
		this.cellFont = cellFont;
	}

	@Override
	public String toString() {
		return new StringBuilder("WrapperCellStyle [backgoundColor=")
				.append(this.backgoundColor).append(", cellFont=")
				.append(this.cellFont).append("]").toString();
	}
}
