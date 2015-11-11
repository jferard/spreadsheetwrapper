package com.github.jferard.spreadsheetwrapper.impl;

import java.util.HashMap;
import java.util.Map;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * @param <T>
 *            the *internal* cell style class
 */
public class CellStyleAccessor<T> {
	Map<String, T> cellStyleByName;
	Map<T, String> nameByCellStyle;

	public CellStyleAccessor() {
		this.cellStyleByName = new HashMap<String, T>();
		this.nameByCellStyle = new HashMap<T, String>();
	}

	public/*@Nullable*/T getCellStyle(final String name) {
		return this.cellStyleByName.get(name);
	}

	public/*@Nullable*/String getName(final T cellStyle) {
		return this.nameByCellStyle.get(cellStyle);
	}

	public void putCellStyle(final String name, final T style) {
		this.cellStyleByName.put(name, style);
		this.nameByCellStyle.put(style, name);
	}

}
