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

public class TestProperties {
	final String extension;
	final SpreadsheetDocumentFactory factory;

	public TestProperties(final String extension,
			final SpreadsheetDocumentFactory factory) {
		this.extension = extension;
		this.factory = factory;
	}

	public String getExtension() {
		return this.extension;
	}

	public SpreadsheetDocumentFactory getFactory() {
		return this.factory;
	}

	public URL getSourceURL() {
		final String sourceURLString = String.format(
				"/VilleMTP_MTP_MonumentsHist.%s", this.extension);
		return this.getClass().getResource(sourceURLString);
	}

}
