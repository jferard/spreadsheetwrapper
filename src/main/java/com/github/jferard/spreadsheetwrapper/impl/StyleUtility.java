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

import java.util.HashMap;
import java.util.Map;

public class StyleUtility {

	/**
	 * @param styleString
	 *            the styleString, format key1:value1;key2:value2
	 * @return
	 */
	public Map<String, String> getPropertiesMap(final String styleString) {
		final Map<String, String> properties = new HashMap<String, String>();
		final String[] styleProps = styleString.split(";");
		for (final String styleProp : styleProps) {
			final String[] entry = styleProp.split(":");
			properties.put(entry[0].trim().toLowerCase(), entry[1].trim());
		}
		return properties;
	}

}
