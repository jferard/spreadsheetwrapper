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

import com.github.jferard.spreadsheetwrapper.CellStyle.Color;

public class Font {
	private boolean bold;
	private boolean italic;
	private int size;
	private final CellStyle.Color color;
	
	public Font(boolean bold, boolean italic, int size, Color color) {
		super();
		this.bold = bold;
		this.italic = italic;
		this.size = size;
		this.color = color;
	}

	public boolean isBold() {
		return this.bold;
	}

	public void setBold(boolean bold) {
		this.bold = bold;
	}

	public boolean isItalic() {
		return this.italic;
	}

	public void setItalic(boolean italic) {
		this.italic = italic;
	}

	public int getSize() {
		return this.size;
	}

	public void setSize(int size) {
		this.size = size;
	}

	public CellStyle.Color getColor() {
		return this.color;
	}

}
