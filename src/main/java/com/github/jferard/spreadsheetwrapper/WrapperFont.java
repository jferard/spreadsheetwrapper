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

/**
 * the class WrapperFont wraps some font attributes : bold/italic, size and
 * color
 *
 */
public class WrapperFont {
	/** true if the font is bold */
	private boolean bold;
	/** true if the font is italic */
	private boolean italic;
	/** size of the font */
	private int size;
	/** color of the font */
	private final WrapperColor wrapperColor;

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
	public WrapperFont(final boolean bold, final boolean italic,
			final int size, final WrapperColor wrapperColor) {
		super();
		this.bold = bold;
		this.italic = italic;
		this.size = size;
		this.wrapperColor = wrapperColor;
	}

	/**
	 * @return the color of the font
	 */
	public WrapperColor getColor() {
		return this.wrapperColor;
	}

	/**
	 * @return the size of the font
	 */
	public int getSize() {
		return this.size;
	}

	/**
	 * @return true if the font is bold
	 */
	public boolean isBold() {
		return this.bold;
	}

	/**
	 * @return true if the font is italic
	 */
	public boolean isItalic() {
		return this.italic;
	}

	/**
	 * @param bold
	 *            true toset the font to bold
	 */
	public void setBold(final boolean bold) {
		this.bold = bold;
	}

	/**
	 * @param italic
	 *            true to set thee font to italic
	 */
	public void setItalic(final boolean italic) {
		this.italic = italic;
	}

	/**
	 * @param size
	 *            size of the font
	 */
	public void setSize(final int size) {
		this.size = size;
	}

}
