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

import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.junit.Assert;
import org.junit.Assume;
import org.junit.Before;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.TestProperties;

public abstract class AbstractSpreadsheetReaderCursorImplTest extends
CursorAbstractTest {

	/** simple logger, static initilization */
	private final Logger logger = Logger.getLogger(this.getClass().getName());

	/** set the test up */
	@Before
	@SuppressWarnings("nullness")
	public void setUp() {
		final SpreadsheetDocumentFactory factory = this.getProperties()
				.getFactory();
		try {
			final URL sourceURL = this.getProperties().getResourceURL();
			Assume.assumeNotNull(sourceURL);

			final InputStream inputStream = sourceURL.openStream();
			final SpreadsheetDocumentReader sdr = factory
					.openForRead(inputStream);
			final SpreadsheetReader sheet = sdr.getSpreadsheet(0);
			this.rowCount = sheet.getRowCount();
			this.colCount = sheet.getCellCount(0);
			this.cursor = sheet.getNewCursor();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		} catch (final IOException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** @return properties for test */
	protected abstract TestProperties getProperties();
}
