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
import java.util.logging.Level;
import java.util.logging.Logger;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;

import com.github.jferard.spreadsheetwrapper.AbstractSpreadsheetEmptyWriterTest;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetTestHelper;
import com.github.jferard.spreadsheetwrapper.TestProperties;

public class XlsJxlEmptyWriterTest extends AbstractSpreadsheetEmptyWriterTest {
	/** logger */
	private Logger logger;

	/** set the test up */
	@Before
	@Override
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		this.logger = Logger.getLogger(this.getClass().getName());
		try {
			final File outputFile = SpreadsheetTestHelper.getOutputFile(
					this.factory, this.getClass().getSimpleName(),
					this.name.getMethodName());
			this.documentWriter = this.factory.create(outputFile);
			this.sheetWriter = this.documentWriter.addSheet(0, "first sheet");
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.INFO, "", e);
			Assert.fail();
		}
	}

	/** tear the test down */
	@After
	@Override
	public void tearDown() {
		try {

			this.documentWriter.save();
			this.documentWriter.close();
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
