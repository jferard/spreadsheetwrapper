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
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.NoSuchElementException;

import org.junit.After;
import org.junit.Assert;
import org.junit.Assume;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;

/**
 * Okay for all wrappers
 *
 */
@SuppressWarnings("PMD")
public abstract class SpreadsheetDocumentWriterTest extends
SpreadsheetDocumentReaderTest {
	/** name of the test */
	@Rule
	public TestName name = new TestName();

	/** the document writer */
	protected SpreadsheetDocumentWriter sdw;

	/** the writer */
	protected SpreadsheetWriter sw;

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
			this.sdw = this.factory.openForWrite(inputStream);
			this.sdr = this.sdw;
			Assert.assertEquals(1, this.sdw.getSheetCount());
			this.sw = this.sdw.getSpreadsheet(0);
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		} catch (final IOException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	/** tear the test down */
	@Override
	@After
	public void tearDown() {
		try {
			final File outputFile = SpreadsheetTestHelper.getOutputFile(this
					.getClass().getSimpleName(), this.name.getMethodName(),
					this.getProperties().getExtension());
			this.sdw.saveAs(outputFile);
			this.sdw.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test
	public void testAddSheetAtIndex0() throws IndexOutOfBoundsException,
	CantInsertElementInSpreadsheetException {
		try {
			this.sdw.addSheet(0, "ok");
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test(expected = IndexOutOfBoundsException.class)
	public final void testAddSheetAtIndex10() throws IndexOutOfBoundsException,
	CantInsertElementInSpreadsheetException {
		this.sdw.addSheet(10, "ok");
	}

	@Test(expected = IndexOutOfBoundsException.class)
	public final void testAddSheetAtNegativeIndex()
			throws IndexOutOfBoundsException,
			CantInsertElementInSpreadsheetException {
		this.sdw.addSheet(-1, "ok");
	}

	@Test
	public final void testAppend1000SheetsAtEnd()
			throws CantInsertElementInSpreadsheetException {
		for (int i = 0; i < 1000; i++) {
			this.sdw.addSheet("new" + i);
		}
		Assert.assertTrue(true);
	}

	@Test
	public void testAppendAdd20SheetAtIndex0()
			throws CantInsertElementInSpreadsheetException {
		try {
			for (int i = 0; i < 20; i++) {
				this.sdw.addSheet(0, "new" + i);
			}
			Assert.assertTrue(true);
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public final void testAppendAndThenGetSheet()
			throws CantInsertElementInSpreadsheetException {
		this.sdw.addSheet("ok");
		try {
			this.sdw.getSpreadsheet("ok");
		} catch (final NoSuchElementException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test
	public final void testAppendSheet()
			throws CantInsertElementInSpreadsheetException {
		this.sdw.addSheet("ok");
	}

	@Test
	public final void testAppendSheetAndThenGetByIndex()
			throws CantInsertElementInSpreadsheetException {
		this.sdw.addSheet("ok");
		try {
			this.sdw.getSpreadsheet(1);
		} catch (final NoSuchElementException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test(expected = IllegalArgumentException.class)
	public final void testAppendSheetWithExistingName()
			throws CantInsertElementInSpreadsheetException {
		this.sdw.addSheet("Feuille1");
	}

	@Override
	protected abstract TestProperties getProperties();
}