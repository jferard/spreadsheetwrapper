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
package com.github.jferard.spreadsheetwrapper.style;

import com.github.jferard.spreadsheetwrapper.Util;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * the class WrapperFont wraps some font attributes : bold/italic, size, color
 * and family
 *
 */
public class WrapperFont {
	/** ARIAL font */
	public static final String ARIAL_NAME = "Arial";
	/** COURIER font */
	public static final String COURIER_NAME = "Courier New";
	/** TAHOMA font */
	public static final String TAHOMA_NAME = "Tahoma";
	/** TIMES font */
	public static final String TIMES_NAME = "Times New Roman";

	/** YES if the font is bold */
	private int bold;
	/** family of the font : one of */
	private/*@Nullable*/String family;
	/** YES if the font is italic */
	private int italic;
	/** size of the font, -1 if no size specified */
	private double size;
	/** color of the font */
	private/*@Nullable*/WrapperColor wrapperColor;

	/**
	 * All fields to default
	 */
	public WrapperFont() {
		this.bold = WrapperCellStyle.DEFAULT;
		this.italic = WrapperCellStyle.DEFAULT;
		this.size = WrapperCellStyle.DEFAULT;
		this.wrapperColor = null;
		this.family = null;
	}

	/**
	 * @param bold
	 *            one of YES, NO, DEFAULT
	 * @param italic
	 *            one of YES, NO, DEFAULT
	 */
	public WrapperFont(final int bold, final int italic) {
		this(bold, italic, WrapperCellStyle.DEFAULT, null, null);
	}

	/**
	 * @param bold
	 *            true if the font is bold
	 * @param italic
	 *            true if the font is italic
	 * @param size
	 *            size of the font
	 * @param wrapperColor
	 *            color of the font
	 */
	public WrapperFont(final int bold, final int italic, final double size,
			final/*@Nullable*/WrapperColor wrapperColor,
			final/*@Nullable*/String family) {
		super();
		this.bold = bold;
		this.italic = italic;
		this.size = size;
		this.wrapperColor = wrapperColor;
		this.family = family;
	}

	/** {@inheritDoc} */
	@Override
	public boolean equals(final/*@Nullable*/Object obj) {
		if (this == obj)
			return true;
		if (!(obj instanceof WrapperFont))
			return false;

		final WrapperFont other = (WrapperFont) obj;
		return this.bold == other.bold && this.italic == other.italic
				&& Util.almostEqual(this.size, other.size)
				&& this.wrapperColor == other.wrapperColor
				&& Util.equal(this.family, other.family);
	}

	/**
	 * @return true if the font is bold
	 */
	public int getBold() {
		return this.bold;
	}

	/**
	 * @return the color of the font
	 */
	public/*@Nullable*/WrapperColor getColor() {
		return this.wrapperColor;
	}

	/**
	 * @return the family of the font
	 */
	public/*@Nullable*/String getFamily() {
		return this.family;
	}

	/**
	 * @return true if the font is italic
	 */
	public int getItalic() {
		return this.italic;
	}

	/**
	 * @return the size of the font
	 */
	public double getSize() {
		return this.size;
	}

	/** {@inheritDoc} */
	@Override
	public int hashCode() {
		final int prime = 31;
		return prime
				* (prime * (prime * (prime + this.bold) + this.italic) + (int) this.size)
				+ Util.hash(this.wrapperColor) + Util.hash(this.family);
	}

	/**
	 * sets font to bold
	 *
	 * @return this (fluent)
	 */
	public WrapperFont setBold() {
		this.bold = WrapperCellStyle.YES;
		return this;
	}

	/**
	 * @param bold
	 *            one of YES, NO, DEFAULT
	 * @return this (fluent)
	 */
	public WrapperFont setBold(final int bold) {
		this.bold = bold;
		return this;
	}

	/**
	 * sets font color
	 *
	 * @param color
	 *            the WrapperColor
	 * @return this (fluent)
	 */
	public WrapperFont setColor(final WrapperColor color) {
		this.wrapperColor = color;
		return this;
	}

	/**
	 * sets font family
	 *
	 * @param family
	 *            the family name
	 * @return this (fluent)
	 */
	public WrapperFont setFamily(final String family) {
		this.family = family;
		return this;
	}

	/**
	 * sets font to italic
	 *
	 * @return this (fluent)
	 */
	public WrapperFont setItalic() {
		this.italic = WrapperCellStyle.YES;
		return this;
	}

	/**
	 * @param italic
	 *            one of YES, NO, DEFAULT
	 * @return this (fluent)
	 */
	public WrapperFont setItalic(final int italic) {
		this.italic = italic;
		return this;
	}

	/**
	 * @param size
	 *            size of the font
	 * @return this (fluent)
	 */
	public WrapperFont setSize(final double size) {
		this.size = size;
		return this;
	}

	/** {@inheritDoc} */
	@Override
	public String toString() {
		return new StringBuilder("WrapperFont [bold=").append(this.bold)
				.append(", italic=").append(this.italic).append(", size=")
				.append(this.size).append(", wrapperColor=")
				.append(this.wrapperColor).append(", family=")
				.append(this.family).append("]").toString();
	}
}
