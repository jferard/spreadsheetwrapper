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

public class WrapperFont {
	private boolean bold;
	private boolean italic;
	private int size;
	private final WrapperColor wrapperColor;

	public WrapperFont(final boolean bold, final boolean italic,
			final int size, final WrapperColor wrapperColor) {
		super();
		this.bold = bold;
		this.italic = italic;
		this.size = size;
		this.wrapperColor = wrapperColor;
	}

	public WrapperColor getColor() {
		return this.wrapperColor;
	}

	public int getSize() {
		return this.size;
	}

	public boolean isBold() {
		return this.bold;
	}

	public boolean isItalic() {
		return this.italic;
	}

	public void setBold(final boolean bold) {
		this.bold = bold;
	}

	public void setItalic(final boolean italic) {
		this.italic = italic;
	}

	public void setSize(final int size) {
		this.size = size;
	}

}
