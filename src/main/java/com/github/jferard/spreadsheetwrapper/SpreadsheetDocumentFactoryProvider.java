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

import java.util.Map;

/**
 * A provider for spreadsheet factories. Usage :
 * SpreadsheetDocumentFactoryProvider sdfp = new
 * SpreadsheetDocumentFactoryProvider(...); SpreadsheetDocumentFactory sdf =
 * sdfp.getFactory("xls"); SpreadsheetDocumentWriter sdw = sdf.create();
 */
public class SpreadsheetDocumentFactoryProvider {
	/** Map (ods|xls) -> factory */
	private final Map<String, SpreadsheetDocumentFactory> factoryByExtension;

	/**
	 * @param factoryByExtension
	 *            Map (ods|xls) -> factory
	 */
	public SpreadsheetDocumentFactoryProvider(
			final Map<String, SpreadsheetDocumentFactory> factoryByExtension) {
		this.factoryByExtension = factoryByExtension;
	}

	/**
	 * @param extension
	 *            xls|ods
	 * @return the factory
	 */
	public SpreadsheetDocumentFactory getFactory(final String extension) {
		if (this.factoryByExtension.containsKey(extension))
			return this.factoryByExtension.get(extension);
		else
			throw new IllegalArgumentException(String.format(
					"Can't instantiate factory for extension %s", extension));
	}
}