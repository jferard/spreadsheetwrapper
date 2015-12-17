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

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * Tiny Util class.
 */
public final class Util {
	/**
	 * @param aDouble
	 * @param otherDouble
	 * @return true if aDouble and otherDouble are equal to 0.001
	 */
	public static boolean almostEqual(final double aDouble,
			final double otherDouble) {
		return Math.abs(aDouble - otherDouble) < 0.001;
	}

	/**
	 * @param object
	 * @param otherObject
	 * @return true if object and otherObject are equal
	 */
	public static boolean equal(final/*@Nullable*/Object object,
			final/*@Nullable*/Object otherObject) {
		return object == null ? otherObject == null : object
				.equals(otherObject);
	}

	/**
	 * @param object
	 *            the objet to hash
	 * @return 0 if object is null, else object.hashCode
	 */
	public static int hash(final/*@Nullable*/Object object) {
		return object == null ? 0 : object.hashCode();
	}

	private Util() {
	}
}
