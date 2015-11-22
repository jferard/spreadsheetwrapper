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
package com.github.jferard.spreadsheetwrapper.xls.jxl;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.junit.After;
import org.junit.Assert;
import org.junit.Assume;
import org.junit.Before;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriterTest;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetTestHelper;
import com.github.jferard.spreadsheetwrapper.TestProperties;

public class XlsJxlDocumentWriterTest extends SpreadsheetDocumentWriterTest {
	/** the logger */
	private final Logger logger = Logger.getLogger(this.getClass().getName());

	/** set the test up */
	@Override
	@Before
	@SuppressWarnings("nullness")
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			final URL sourceURL = this.getProperties().getSourceURL();
			Assume.assumeNotNull(sourceURL);

			final InputStream inputStream = sourceURL.openStream();
			final File outputFile = SpreadsheetTestHelper.getOutputFile(this
					.getClass().getSimpleName(), this.name.getMethodName(),
					this.getProperties().getExtension());
			final OutputStream outputStream = new FileOutputStream(outputFile);
			this.sdw = this.factory.openForWrite(inputStream, outputStream);
			this.documentReader = this.sdw;
			Assert.assertEquals(1, this.sdw.getSheetCount());
			this.sw = this.sdw.getSpreadsheet(0);
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.INFO, "", e);
			Assert.fail();
		} catch (final IOException e) {
			this.logger.log(Level.INFO, "", e);
			Assert.fail();
		}
	}

	/** tear the test down */
	@Override
	@After
	public void tearDown() {
		try {
			this.sdw.save();
			this.sdw.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.INFO, "", e);
			Assert.fail();
		}
	}

	/** {@inheritDoc} */
	@Override
	protected TestProperties getProperties() {
		return XlsJxlTestProperties.getProperties();
	}
}
