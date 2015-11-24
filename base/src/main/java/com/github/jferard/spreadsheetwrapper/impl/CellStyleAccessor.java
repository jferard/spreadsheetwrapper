package com.github.jferard.spreadsheetwrapper.impl;

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
