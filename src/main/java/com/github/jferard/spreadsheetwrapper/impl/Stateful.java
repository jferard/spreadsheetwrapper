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

import org.jopendocument.dom.spreadsheet.Sheet;

public class Stateful<T> {
	protected final T value;
	private boolean isNew;

	public static <S> Stateful<S> createNew(S document) {
		return new Stateful<S>(document, true);
	}
	
	public static <S> Stateful<S> createInitialized(S document) {
		return new Stateful<S>(document, false);
	}
	
	protected Stateful(T value, boolean isNew) {
		this.value = value;
		this.isNew = isNew;
	}

	public T getValue() {
		return this.value;
	}

	public boolean isNew() {
		return this.isNew;
	}

	public void setInitialized() {
		this.isNew = false;
	}
}
