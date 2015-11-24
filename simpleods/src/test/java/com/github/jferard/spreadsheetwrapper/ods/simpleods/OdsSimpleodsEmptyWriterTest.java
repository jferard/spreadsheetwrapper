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
package com.github.jferard.spreadsheetwrapper.ods.simpleods;

import java.io.File;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;

import com.github.jferard.spreadsheetwrapper.AbstractSpreadsheetEmptyWriterTest;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetTestHelper;
import com.github.jferard.spreadsheetwrapper.TestProperties;

public class OdsSimpleodsEmptyWriterTest extends
AbstractSpreadsheetEmptyWriterTest {
	/** set the test up */
	@Before
	@Override
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			final File outputFile = SpreadsheetTestHelper.getOutputFile(this
					.getClass().getSimpleName(), this.name.getMethodName(),
					this.getProperties().getExtension());
			this.documentWriter = this.factory.create(outputFile);
			this.sheetWriter = this.documentWriter.addSheet(0, "first sheet");
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
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
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Override
	protected TestProperties getProperties() {
		return OdsSimpleodsTestProperties.getProperties();
	}
}
