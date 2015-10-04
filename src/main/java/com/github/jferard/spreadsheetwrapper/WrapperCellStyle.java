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


public class WrapperCellStyle {
	private WrapperColor backgoundColor;
	private WrapperFont cellFont;

	public WrapperCellStyle(final WrapperColor wrapperColor,
			final WrapperFont cellFont) {
		super();
		this.backgoundColor = wrapperColor;
		this.cellFont = cellFont;
	}

	public WrapperColor getBackgroundColor() {
		return this.backgoundColor;
	}

	public WrapperFont getCellFont() {
		return this.cellFont;
	}

	public void setBackgroundColor(final WrapperColor wrapperColor) {
		this.backgoundColor = wrapperColor;
	}

	public void setCellFont(final WrapperFont cellFont) {
		this.cellFont = cellFont;
	}
}
