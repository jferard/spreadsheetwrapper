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

import java.net.URL;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * The class TestProperties is a container for basic properties of tests :
 * extension of file (xls/ods) and factory used
 *
 */
public class TestProperties {
	/** the factory to use */
	private final SpreadsheetDocumentFactory factory;

	/**
	 * @param factory
	 *            the factory to use
	 */
	public TestProperties(final SpreadsheetDocumentFactory factory) {
		this.factory = factory;
	}

	/**
	 * @return the factory to use
	 */
	public SpreadsheetDocumentFactory getFactory() {
		return this.factory;
	}

	/**
	 * @return the URL of the resource
	 */
	public/*@Nullable*/URL getResourceURL() {
		final String sourceURLString = String.format(
				"/VilleMTP_MTP_MonumentsHist.%s", this.factory.getExtension());
		return this.getClass().getResource(sourceURLString);
	}
}
