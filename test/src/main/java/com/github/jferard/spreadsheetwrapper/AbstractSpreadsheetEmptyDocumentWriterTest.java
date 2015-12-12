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
public abstract class AbstractSpreadsheetEmptyDocumentWriterTest extends
AbstractSpreadsheetEmptyDocumentReaderTest {
	/** the test name */
	@Rule
	public TestName name = new TestName();

	/** the document writer */
	protected SpreadsheetDocumentWriter documentWriter;

	/** a sheet writer */
	protected SpreadsheetWriter sheetWriter;

	/** set the test up */
	@Override
	@Before
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			final File outputFile = SpreadsheetTestHelper.getOutputFile(
					this.factory, this.getClass().getSimpleName(),
					this.name.getMethodName());
			this.documentWriter = this.factory.create(outputFile);
			this.documentReader = this.documentWriter;
		} catch (final SpreadsheetException e) {
			Assert.fail();
		}
	}

	/** tear the test down */
	@Override
	@After
	public void tearDown() {
		try {
			this.documentWriter.save();
			this.documentWriter.close();
		} catch (final SpreadsheetException e) {
			Assert.fail();
		}
	}

	/** test if there is no sheet and then add one */
	@Test
	public final void testAddSheet()
			throws CantInsertElementInSpreadsheetException {
		final String SHEET_NAME = "ok";
		Assert.assertEquals(0, this.documentWriter.getSheetCount());
		Assert.assertEquals(Collections.emptyList(),
				this.documentWriter.getSheetNames());
		this.documentWriter.addSheet(SHEET_NAME);
		Assert.assertEquals(1, this.documentWriter.getSheetCount());
		Assert.assertEquals(Arrays.asList(SHEET_NAME),
				this.documentWriter.getSheetNames());
		Assert.assertNotNull(this.documentWriter.getSpreadsheet(0));
		Assert.assertNotNull(this.documentWriter.getSpreadsheet(SHEET_NAME));
	}

	/** there is no sheet a beginning */
	@Test(expected = IndexOutOfBoundsException.class)
	public final void testGetNonExistingSheet() {
		this.documentWriter.getSpreadsheet(0);
	}
}