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

import java.io.File;
import java.util.Arrays;
import java.util.Collections;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;

/**
 * Okay for all wrappers
 *
 */
@SuppressWarnings("PMD")
public abstract class SpreadsheetEmptyDocumentWriterTest extends
		AbstractSpreadsheetEmptyDocumentReaderTest {
	/** the test name */
	@Rule
	public TestName name = new TestName();
	protected SpreadsheetDocumentWriter sdw;

	protected SpreadsheetWriter sw;

	/** set the test up */
	@Override
	@Before
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			final File outputFile = SpreadsheetTestHelper.getOutputFile(this
					.getClass().getSimpleName(), this.name.getMethodName(),
					this.getProperties().getExtension());
			this.sdw = this.factory.create(outputFile);
			this.documentReader = this.sdw;
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
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
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test
	public final void testAddSheet()
			throws CantInsertElementInSpreadsheetException {
		Assert.assertEquals(0, this.sdw.getSheetCount());
		Assert.assertEquals(Collections.emptyList(), this.sdw.getSheetNames());
		this.sdw.addSheet("ok");
		Assert.assertEquals(1, this.sdw.getSheetCount());
		Assert.assertEquals(Arrays.asList("ok"), this.sdw.getSheetNames());
		Assert.assertNotNull(this.sdw.getSpreadsheet(0));
		Assert.assertNotNull(this.sdw.getSpreadsheet("ok"));
	}

	@Test(expected = IndexOutOfBoundsException.class)
	public final void testGetNonExistingSheet() {
		this.sdw.getSpreadsheet(0);
	}
}