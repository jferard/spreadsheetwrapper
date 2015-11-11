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
 * the class WrapperFont wraps some font attributes : bold/italic, size and
 * color
 *
 */
public class WrapperFont {
	/** true if the font is bold */
	private int bold;
	/** true if the font is italic */
	private int italic;
	/** size of the font, -1 if no size specified */
	private int size;
	/** color of the font */
	private/*@Nullable*/WrapperColor wrapperColor;

	public WrapperFont() {
		this.bold = WrapperCellStyle.DEFAULT;
		this.italic = WrapperCellStyle.DEFAULT;
		this.size = WrapperCellStyle.DEFAULT;
		this.wrapperColor = null;
	}

	public WrapperFont(final int bold, final int italic) {
		this(bold, italic, WrapperCellStyle.DEFAULT, null);
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
	public WrapperFont(final int bold, final int italic, final int size,
			final/*@Nullable*/WrapperColor wrapperColor) {
		super();
		this.bold = bold;
		this.italic = italic;
		this.size = size;
		this.wrapperColor = wrapperColor;
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
	 * @return true if the font is italic
	 */
	public int getItalic() {
		return this.italic;
	}

	/**
	 * @return the size of the font
	 */
	public int getSize() {
		return this.size;
	}

	public WrapperFont setBold() {
		this.bold = WrapperCellStyle.YES;
		return this;
	}

	/**
	 * @param bold
	 *            true toset the font to bold
	 */
	public WrapperFont setBold(final int bold) {
		this.bold = bold;
		return this;
	}

	public WrapperFont setColor(final WrapperColor color) {
		this.wrapperColor = color;
		return this;
	}

	/**
	 * @param italic
	 *            true to set thee font to italic
	 */
	public WrapperFont setItalic() {
		this.italic = WrapperCellStyle.YES;
		return this;
	}

	/**
	 * @param italic
	 *            true to set thee font to italic
	 */
	public WrapperFont setItalic(final int italic) {
		this.italic = italic;
		return this;
	}

	/**
	 * @param size
	 *            size of the font
	 */
	public WrapperFont setSize(final int size) {
		this.size = size;
		return this;
	}
}
