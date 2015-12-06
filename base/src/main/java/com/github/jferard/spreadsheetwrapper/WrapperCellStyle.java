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

/**
 * @author Julien
 *
 */
public class WrapperCellStyle {
	/** attribute : use the default value */
	public static final int DEFAULT = -1;

	/** empty style */
	public static final WrapperCellStyle EMPTY = new WrapperCellStyle(null, -1.0,
			new WrapperFont());
	/** attribute : do not set */
	public static final int NO = 0; // NOPMD by Julien on 20/11/15 21:04
	/** attribute : set */
	public static final int YES = 1;

	/** the backgroundColor of the cell */
	private/*@Nullable*/WrapperColor backgoundColor;

	/** the border size of the cell */
	private double borderLineWidth;

	/** the text font */
	private/*@Nullable*/WrapperFont cellFont;

	/**
	 * @param backgroundColor
	 *            the background color of the cell
	 * @param cellFont
	 *            the font
	 */
	public WrapperCellStyle(final/*@Nullable*/WrapperColor backgroundColor,
			final double borderLineWidth, final/*@Nullable*/WrapperFont cellFont) {
		super();
		this.backgoundColor = backgroundColor;
		this.borderLineWidth = borderLineWidth;
		this.cellFont = cellFont;
	}

	/** {@inheritDoc} */
	@Override
	public boolean equals(final/*@Nullable*/Object obj) {
		if (this == obj)
			return true;
		if (!(obj instanceof WrapperCellStyle))
			return false;

		final WrapperCellStyle other = (WrapperCellStyle) obj;
		return this.backgoundColor == other.backgoundColor
				&& Util.almostEqual(this.borderLineWidth, other.borderLineWidth)
				&& Util.equal(this.cellFont, other.cellFont);
	}

	/**
	 * @return the background color of the cell
	 */
	public/*@Nullable*/WrapperColor getBackgroundColor() {
		return this.backgoundColor;
	}
	
	/**
	 * @return the background color of the cell
	 */
	public double getBorderLineWidth() {
		return this.borderLineWidth;
	}

	/**
	 * @return the text font in the cell
	 */
	public/*@Nullable*/WrapperFont getCellFont() {
		return this.cellFont;
	}

	/** {@inheritDoc} */
	@Override
	public int hashCode() {
		final int prime = 31;
		return prime * (prime + Util.hash(this.backgoundColor))
				+ Util.hash(this.cellFont);
	}

	/**
	 * @param wrapperColor
	 *            the background color to set
	 * @return this for fluent style
	 */
	public WrapperCellStyle setBackgroundColor(final WrapperColor wrapperColor) {
		this.backgoundColor = wrapperColor;
		return this;
	}
	
	/**
	 * @param borderLineWidth
	 *            the width of the line (border)
	 * @return this for fluent style
	 */
	public WrapperCellStyle setBackgroundColor(final double borderLineWidth) {
		this.borderLineWidth = borderLineWidth;
		return this;
	}

	/**
	 * @param cellFont
	 *            the font to set
	 * @return this for fluent style
	 */
	public WrapperCellStyle setCellFont(final WrapperFont cellFont) {
		this.cellFont = cellFont;
		return this;
	}

	/** {@inheritDoc} */
	@Override
	public String toString() {
		return new StringBuilder("WrapperCellStyle [backgoundColor=")
				.append(this.backgoundColor).append(", borderLineWidth=")
				.append(this.borderLineWidth).append(", cellFont=")
				.append(this.cellFont).append("]").toString();
	}
}
