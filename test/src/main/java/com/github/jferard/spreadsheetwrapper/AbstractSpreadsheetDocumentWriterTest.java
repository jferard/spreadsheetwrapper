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
import java.util.logging.Level;
import java.util.logging.Logger;

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
public abstract class AbstractSpreadsheetDocumentWriterTest extends
AbstractSpreadsheetDocumentReaderTest {
	/** name of the test */
	@Rule
	public TestName name = new TestName();

	/** the document writer */
	protected SpreadsheetDocumentWriter documentWriter;

	/** the writer */
	protected SpreadsheetWriter sheetWriter;

	/** logger, static initialization */
	private final Logger logger = Logger.getLogger(this.getClass().getName());
	
	/** set the test up */
	@Override
	@Before
	@SuppressWarnings("nullness")
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			final URL sourceURL = this.getProperties().getResourceURL();
			Assume.assumeNotNull(sourceURL);

			final InputStream inputStream = sourceURL.openStream();
			this.documentWriter = this.factory.openForWrite(inputStream);
			this.documentReader = this.documentWriter;
			Assert.assertEquals(1, this.documentWriter.getSheetCount());
			this.sheetWriter = this.documentWriter.getSpreadsheet(0);
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
			final File outputFile = SpreadsheetTestHelper.getOutputFile(this
					.getClass().getSimpleName(), this.name.getMethodName(),
					this.getProperties().getExtension());
			this.documentWriter.saveAs(outputFile);
			this.documentWriter.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** add a sheet "ok" at index 0 */
	@Test
	public void testAddSheetAtIndex0() throws IndexOutOfBoundsException,
	CantInsertElementInSpreadsheetException {
		try {
			this.documentWriter.addSheet(0, "ok");
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	/** add a sheet "ok" at index 10 : not ok because there is no other sheet */
	@Test(expected = IndexOutOfBoundsException.class)
	public final void testAddSheetAtIndex10() throws IndexOutOfBoundsException,
	CantInsertElementInSpreadsheetException {
		this.documentWriter.addSheet(10, "ok");
	}

	/** add a sheet "ok" at index -1 : not ok */
	@Test(expected = IndexOutOfBoundsException.class)
	public final void testAddSheetAtNegativeIndex()
			throws IndexOutOfBoundsException,
			CantInsertElementInSpreadsheetException {
		this.documentWriter.addSheet(-1, "ok");
	}

	/** add 1000 sheets */
	@Test
	public final void testAppend1000SheetsAtEnd()
			throws CantInsertElementInSpreadsheetException {
		for (int i = 0; i < 1000; i++) {
			this.documentWriter.addSheet("new" + i);
		}
		Assert.assertTrue(true);
	}

	/** add 20 sheets at index 0*/
	@Test
	public void testAppendAdd20SheetAtIndex0()
			throws CantInsertElementInSpreadsheetException {
		try {
			for (int i = 0; i < 20; i++) {
				this.documentWriter.addSheet(0, "new" + i);
			}
			Assert.assertTrue(true);
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	/** add sheet "ok" and get it */
	@Test
	public final void testAppendAndThenGetSheet()
			throws CantInsertElementInSpreadsheetException {
		this.documentWriter.addSheet("ok");
		try {
			this.documentWriter.getSpreadsheet("ok");
		} catch (final NoSuchElementException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** add sheet "ok" */
	@Test
	public final void testAppendSheet()
			throws CantInsertElementInSpreadsheetException {
		this.documentWriter.addSheet("ok");
	}

	/** add sheet "ok" : it's sheet number 1 */
	@Test
	public final void testAppendSheetAndThenGetByIndex()
			throws CantInsertElementInSpreadsheetException {
		this.documentWriter.addSheet("ok");
		try {
			this.documentWriter.getSpreadsheet(1);
		} catch (final NoSuchElementException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** can't append a sheet with the same name as sheet 0 */
	@Test(expected = IllegalArgumentException.class)
	public final void testAppendSheetWithExistingName()
			throws CantInsertElementInSpreadsheetException {
		this.documentWriter.addSheet("Feuille1");
	}
}