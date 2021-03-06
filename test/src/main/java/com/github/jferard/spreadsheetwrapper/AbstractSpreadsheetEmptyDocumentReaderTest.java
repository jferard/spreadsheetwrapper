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

import java.util.Collections;
import java.util.List;
import java.util.NoSuchElementException;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

@SuppressWarnings("unused")
public abstract class AbstractSpreadsheetEmptyDocumentReaderTest {

	/** the document reader */
	protected SpreadsheetDocumentReader documentReader;

	/** the factory */
	protected SpreadsheetDocumentFactory factory;

	/** set the test up */
	@Before
	public abstract void setUp();

	/** tear the test down */
	@After
	public abstract void tearDown();

	/** On creation, no sheet in the document */
	@Test(expected = IndexOutOfBoundsException.class)
	public final void testGetSheetAtIndex0() {
		final SpreadsheetReader sheetReader = this.documentReader
				.getSpreadsheet(0);
	}

	/** On creation, no sheet in the document (confirmation 1) */
	@Test(expected = IndexOutOfBoundsException.class)
	public final void testGetSheetAtIndex1() throws SpreadsheetException {
		final SpreadsheetReader sheetReader = this.documentReader
				.getSpreadsheet(1);
	}

	/** On creation, no sheet in the document (confirmation -1) */
	@Test(expected = IndexOutOfBoundsException.class)
	public final void testGetSheetAtNegativeIndex() throws SpreadsheetException {
		final SpreadsheetReader sheetReader = this.documentReader
				.getSpreadsheet(-1);
	}

	/** On creation, no sheet in the document (confirmation "Feuille1") */
	@Test(expected = NoSuchElementException.class)
	public final void testGetSheetByName() {
		final SpreadsheetReader sheetReader = this.documentReader
				.getSpreadsheet("Feuille1");
	}

	/** On creation, no sheet in the document (confirmation "qwerty") */
	@Test(expected = NoSuchElementException.class)
	public final void testGetSheetByWeirdName() throws SpreadsheetException {
		final SpreadsheetReader sheetReader = this.documentReader
				.getSpreadsheet("qwerty");
	}

	/** On creation, no sheet in the document. No cursor though. */
	@Test(expected = IndexOutOfBoundsException.class)
	public final void testGetSheetCursorAtIndex0() throws SpreadsheetException {
		final SpreadsheetReaderCursor cursor = this.documentReader
				.getNewCursorByIndex(0);
	}

	/**
	 * On creation, no sheet in the document. No cursor though (confirmation
	 * "Feuille1").
	 */
	@Test(expected = NoSuchElementException.class)
	public final void testGetSheetCursorsByName() throws SpreadsheetException {
		final SpreadsheetReaderCursor cursor = this.documentReader
				.getNewCursorByName("Feuille1");
	}

	/** On creation, no sheet in the document (confirmation : empty list). */
	@Test
	public final void testGetSheetNames() {
		final List<String> names = this.documentReader.getSheetNames();
		Assert.assertEquals(Collections.emptyList(), names);
	}

	/** On creation, no sheet in the document (confirmation : empty list). */
	@Test(expected = NoSuchElementException.class)
	public final void testGetSheetThatDoesntExist() throws SpreadsheetException {
		final SpreadsheetReader sheetReader = this.documentReader
				.getSpreadsheet("test");
	}

	/** get the properties */
	protected abstract TestProperties getProperties();
}