/*******************************************************************************
 *     SpreadsheetWrapper - An abstraction layer over the API for Excel or Calc
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
import java.util.logging.Logger;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetTest;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterTest;

public class XlsJxlWriterTest extends SpreadsheetWriterTest {
	/** set the test up */
	@Before
	@Override
	public void setUp() {
		this.factory = this.getFactory();
		try {
			final URL resourceURL = this.getClass().getResource(
					String.format("/VilleMTP_MTP_MonumentsHist.%s",
							this.getExtension()));
			final InputStream inputStream = resourceURL.openStream();
			final File outputFile = SpreadsheetTest.getOutputFile(this
					.getClass().getSimpleName(), this.name.getMethodName(),
					this.getExtension());
			final OutputStream outputStream = new FileOutputStream(outputFile);
			this.sdw = this.factory.openForWrite(inputStream, outputStream);
			this.sdr = this.sdw;
			this.sw = this.sdw.getSpreadsheet(0);
			this.sr = this.sw;
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		} catch (final IOException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	/** tear the test down */
	@After
	@Override
	public void tearDown() {
		try {
			this.sdw.save();
			this.sdw.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Override
	protected String getExtension() {
		return "xls";
	}

	@Override
	protected SpreadsheetDocumentFactory getFactory() {
		return new XlsJxlDocumentFactory(Logger.getGlobal());
	}
}
