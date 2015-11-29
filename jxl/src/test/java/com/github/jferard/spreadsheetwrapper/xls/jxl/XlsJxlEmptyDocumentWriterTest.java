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
import java.io.OutputStream;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;

import com.github.jferard.spreadsheetwrapper.AbstractSpreadsheetEmptyDocumentWriterTest;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetTestHelper;
import com.github.jferard.spreadsheetwrapper.TestProperties;

public class XlsJxlEmptyDocumentWriterTest extends // NOPMD by Julien on
		// 27/11/15 20:38
		AbstractSpreadsheetEmptyDocumentWriterTest {
	/** the logger */
	private final Logger logger = Logger.getLogger(this.getClass().getName());

	/** set the test up */
	@Override
	@Before
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			final File outputFile = SpreadsheetTestHelper.getOutputFile(
					this.factory, this.getClass().getSimpleName(),
					this.name.getMethodName());
			final OutputStream outputStream = new FileOutputStream(outputFile);
			this.documentWriter = this.factory.create(outputStream);
			this.documentReader = this.documentWriter;
			Assert.assertEquals(0, this.documentWriter.getSheetCount());
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		} catch (final IOException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** tear the test down */
	@Override
	@After
	public void tearDown() {
		try {
			this.documentWriter.addSheet("for save");
			this.documentWriter.save();
			this.documentWriter.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** {@inheritDoc} */
	@Override
	protected TestProperties getProperties() {
		return XlsJxlTestProperties.getProperties();
	}
}
