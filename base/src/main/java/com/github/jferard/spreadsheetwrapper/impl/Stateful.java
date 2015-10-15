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

/**
 * The class Stateful is used to keep a trace of the state, new or initialized,
 * of an object. See Optional in Guava or Java8 for examples
 *
 * @param <T>
 *            the wrapped object
 */
public class Stateful<T> {
	/**
	 * @param object
	 *            the object that is already initialized
	 * @return the stateful object
	 */
	public static <S> Stateful<S> createInitialized(final S object) {
		return new Stateful<S>(object, false);
	}

	/**
	 * @param object
	 *            the object that is not initialized (new)
	 * @return the stateful object
	 */
	public static <S> Stateful<S> createNew(final S object) {
		return new Stateful<S>(object, true);
	}

	/** true if the object is not initialized */
	private boolean newObject;

	/** the object */
	protected final T object;

	/**
	 * @param object
	 *            the object
	 * @param newObject
	 *            true if the object is initialized
	 */
	protected Stateful(final T object, final boolean newObject) {
		this.object = object;
		this.newObject = newObject;
	}

	/**
	 * @return the object
	 */
	public T getObject() {
		return this.object;
	}

	/**
	 * @return true if not initialized
	 */
	public boolean isNew() {
		return this.newObject;
	}

	/**
	 * sets the object to initialized
	 */
	public void setInitialized() {
		this.newObject = false;
	}
}
