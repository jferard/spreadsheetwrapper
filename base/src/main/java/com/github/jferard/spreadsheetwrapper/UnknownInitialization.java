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
 * The class UnknownInitialization is used to keep a trace of the state, new or initialized,
 * of an object. See Optional in Guava or Java8 for examples
 *
 * @param <T>
 *            the wrapped object
 */
public class UnknownInitialization<T> {
	/**
	 * @param object
	 *            the object that is already initialized
	 * @return the stateful object
	 */
	public static <S> UnknownInitialization<S> createInitialized(final S object) {
		return new UnknownInitialization<S>(object, true);
	}

	/**
	 * @param object
	 *            the object that is not initialized (new)
	 * @return the stateful object
	 */
	public static <S> UnknownInitialization<S> createUninitialized(final S object) {
		return new UnknownInitialization<S>(object, false);
	}

	/** true if the object is not initialized */
	private boolean initialized;

	/** the object */
	protected final T object;

	/**
	 * @param object
	 *            the object
	 * @param initialized
	 *            true if the object is initialized
	 */
	protected UnknownInitialization(final T object, final boolean initialized) {
		this.object = object;
		this.initialized = initialized;
	}

	/**
	 * @return the object
	 */
	public T getObject() {
		return this.object;
	}

	/**
	 * @return true if not initialized
	 **/
	public boolean isInitialized() {
		return this.initialized;
	}

	/**
	 * sets the object to initialized
	 */
	public void setInitialized() {
		this.initialized = false;
	}
}
