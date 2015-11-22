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

import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.Arrays;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.junit.After;
import org.junit.Assert;
import org.junit.Assume;
import org.junit.Before;
import org.junit.Test;

/**
 * Okay for all wrappers. No exception
 *
 */
@SuppressWarnings("unused")
public abstract class AbstractSpreadsheetDocumentReaderTest {
	/** the col count of the test sheet */
	private static final int COL_COUNT = 7;

	/** the row count of the test sheet */
	private static final int ROW_COUNT = 117;

	/** the sheet name */
	private static final String SHEET_NAME = "Feuille1";

	/** logger, static initialization */
	private final Logger logger = Logger.getLogger(this.getClass().getName());
	/** the document reader */
	protected SpreadsheetDocumentReader documentReader;

	/** the factory */
	protected SpreadsheetDocumentFactory factory;

	/** set the test up */
	@Before
	@SuppressWarnings("nullness")
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			final URL sourceURL = this.getProperties().getSourceURL();
			Assume.assumeNotNull(sourceURL);

			final InputStream inputStream = sourceURL.openStream();
			this.documentReader = this.factory.openForRead(inputStream);
			Assert.assertEquals(1, this.documentReader.getSheetCount());
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		} catch (final IOException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** clean ater the tests */
	@After
	public void tearDown() {
		try {
			this.documentReader.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** must return one sheet name */
	@Test
	public final void testGetSheetNames() {
		final List<String> names = this.documentReader.getSheetNames();
		Assert.assertEquals(
				Arrays.asList(AbstractSpreadsheetDocumentReaderTest.SHEET_NAME),
				names);
	}

	/** the sheet -1 does not exist */
	@Test(expected = IndexOutOfBoundsException.class)
	public final void testGetSheetNumberMinusOneThrowsIndexOutOfBounds()
			throws SpreadsheetException {
		final SpreadsheetReader sheetReader = this.documentReader
				.getSpreadsheet(-1);
	}

	/** the sheet 1 does not exist */
	@Test(expected = IndexOutOfBoundsException.class)
	public final void testGetSheetNumberOneThrowsIndexOutOfBounds()
			throws SpreadsheetException {
		final SpreadsheetReader sheetReader = this.documentReader
				.getSpreadsheet(1);
	}

	/** get twice the sheet 0 */
	@Test
	public final void testGetSheetTwice() throws SpreadsheetException {
		SpreadsheetReader sheetReader = this.documentReader
				.getSpreadsheet(AbstractSpreadsheetDocumentReaderTest.SHEET_NAME);
		sheetReader = this.documentReader
				.getSpreadsheet(AbstractSpreadsheetDocumentReaderTest.SHEET_NAME);
	}

	/** get the sheet by index 0 and row and column counts */
	@Test
	public final void testSheetByIndex() {
		try {
			SpreadsheetReader sheetReader = this.documentReader
					.getSpreadsheet(0);
			sheetReader = this.documentReader.getSpreadsheet(0);
			Assert.assertEquals(
					AbstractSpreadsheetDocumentReaderTest.SHEET_NAME,
					sheetReader.getName());
			Assert.assertEquals(
					AbstractSpreadsheetDocumentReaderTest.ROW_COUNT,
					sheetReader.getRowCount());
			Assert.assertEquals(
					AbstractSpreadsheetDocumentReaderTest.COL_COUNT,
					sheetReader.getCellCount(0));
		} catch (final NoSuchElementException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** compares the sheet by index and by name */
	@Test
	public final void testSheetByIndexAndName() {
		try {
			final SpreadsheetReader sheetReader1 = this.documentReader
					.getSpreadsheet(AbstractSpreadsheetDocumentReaderTest.SHEET_NAME);
			final SpreadsheetReader sheetReader2 = this.documentReader
					.getSpreadsheet(0); // use
			// Accessor
			Assert.assertEquals(sheetReader1, sheetReader2);
		} catch (final NoSuchElementException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** get the sheet by name and row and column counts */
	@Test
	public final void testSheetByName() {
		try {
			final SpreadsheetReader sr = this.documentReader
					.getSpreadsheet(AbstractSpreadsheetDocumentReaderTest.SHEET_NAME);
			Assert.assertEquals(
					AbstractSpreadsheetDocumentReaderTest.SHEET_NAME,
					sr.getName());
			Assert.assertEquals(
					AbstractSpreadsheetDocumentReaderTest.ROW_COUNT,
					sr.getRowCount());
			Assert.assertEquals(
					AbstractSpreadsheetDocumentReaderTest.COL_COUNT,
					sr.getCellCount(0));
		} catch (final NoSuchElementException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** get the sheet by non existing name : must throw an exception */
	@Test(expected = NoSuchElementException.class)
	public final void testSheetDoesntExist() throws SpreadsheetException {
		final SpreadsheetReader sheetReader = this.documentReader
				.getSpreadsheet("test");
	}

	/** get the cursors */
	@Test
	public final void testSheetGetCursors() {
		try {
			final SpreadsheetReaderCursor cursor1 = this.documentReader
					.getNewCursorByIndex(0);
			final SpreadsheetReaderCursor cursor2 = this.documentReader
					.getNewCursorByName(AbstractSpreadsheetDocumentReaderTest.SHEET_NAME);
		} catch (final SpreadsheetException e) {
			Assert.fail();
		}
	}

	/** get properties for the test */
	protected abstract TestProperties getProperties();
}