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
package com.github.jferard.spreadsheetwrapper.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/
/*>>> import org.checkerframework.checker.nullness.qual.NonNull;*/
/*>>> import org.checkerframework.checker.nullness.qual.EnsuresNonNullIf;*/

/**
 * An Accessor stores elements that will accessed by index or by name.
 *
 * @param <T>
 *            the type of the elements
 */
public class Accessor</*@NonNull*/T> {
	/** the elements List */
	private final List<T> elementByIndex;
	/** the elements map : name->element */
	private final Map<String, T> elementByName;

	/**
	 * Creates a new Accessor
	 */
	public Accessor() {
		this.elementByIndex = new ArrayList<T>();
		this.elementByName = new HashMap<String, T>();
	}

	/**
	 * @param index
	 *            of the element to look at
	 * @return the element if it exists (ie hasByName(elementName) == true),
	 *         null otherwise
	 */
	public T getByIndex(final int index) {
		return this.elementByIndex.get(index);
	}

	/**
	 * @param elementName
	 *            the element to look at
	 * @return the element if it exists (ie hasByName(elementName) == true),
	 *         null otherwise
	 */
	public/*@Nullable*/T getByName(final String elementName) {
		return this.elementByName.get(elementName);
	}

	/**
	 * @param index
	 *            of the element to look at
	 * @return true if the element exists
	 */
	/*>>> @EnsuresNonNullIf(expression="getByIndex(#1)", result=true)*/
	public boolean hasByIndex(final int index) {
		return 0 <= index && index < this.elementByIndex.size();
	}

	/**
	 * @param elementName
	 *            the element to look at
	 * @return true if the element exists
	 */
	/*>>> @EnsuresNonNullIf(expression="getByName(#1)", result=true)*/
	public boolean hasByName(final String elementName) {
		return this.elementByName.containsKey(elementName);
	}

	/**
	 * Puts an element into accessor
	 *
	 * @param elementName
	 *            the name of the element
	 * @param index
	 *            the index of the element
	 * @param element
	 *            the element to put
	 */
	public void put(final String elementName, final Integer index,
			final T element) {
		this.elementByName.put(elementName, element);
		this.elementByIndex.add(index, element);
	}

	/**
	 * Puts an element into accessor, index = accessor.size();
	 *
	 * @param elementName
	 *            the name of the element
	 * @param element
	 *            the element to put
	 */
	public void putAtEnd(final String elementName, final T element) {
		this.elementByIndex.add(element);
		this.elementByName.put(elementName, element);
	}

	/* {@inheritDoc} */
	@Override
	public String toString() {
		assert this.elementByName.size() == this.elementByIndex.size();
		final StringBuilder stringBuilder = new StringBuilder("Accessor[");
		final ListIterator<T> iterator = this.elementByIndex.listIterator();
		while (iterator.hasNext()) {
			final int index = iterator.nextIndex();
			final T element = iterator.next();
			String name = null;
			for (final Map.Entry<String, T> entry : this.elementByName
					.entrySet()) {
				if (entry.getValue().equals(element)) {
					name = entry.getKey();
					break;
				}
			}
			stringBuilder.append(index).append('&').append(name).append("->")
					.append(element).append(',');
		}
		stringBuilder.deleteCharAt(stringBuilder.length() - 1).append(']');
		return stringBuilder.toString();
	}
}
