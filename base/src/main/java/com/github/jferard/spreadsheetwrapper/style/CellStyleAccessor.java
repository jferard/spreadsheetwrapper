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

import java.util.HashMap;
import java.util.Map;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * @param <T>
 *            the *internal* cell style class
 */
/**
 * @author Julien
 *
 * @param <T>
 */
public class CellStyleAccessor<T> {
	/** name -> cellStyle:T */
	private final Map<String, T> cellStyleByName;
	/** cellStyle:T -> name */
	private final Map<T, String> nameByCellStyle;

	/**
	 * Default constructor
	 */
	public CellStyleAccessor() {
		this.cellStyleByName = new HashMap<String, T>();
		this.nameByCellStyle = new HashMap<T, String>();
	}

	/**
	 * @param name
	 *            name of the style
	 * @return the style, null if none
	 */
	public/*@Nullable*/T getCellStyle(final String name) {
		return this.cellStyleByName.get(name);
	}

	/**
	 * @param cellStyle
	 *            the style
	 * @return the name of the style, null if none
	 */
	public/*@Nullable*/String getName(final T cellStyle) {
		return this.nameByCellStyle.get(cellStyle);
	}

	/**
	 * Adds or update a style
	 *
	 * @param name
	 *            the name
	 * @param style
	 *            the style
	 */
	public void putCellStyle(final String name, final T style) {
		this.cellStyleByName.put(name, style);
		this.nameByCellStyle.put(style, name);
	}
}
