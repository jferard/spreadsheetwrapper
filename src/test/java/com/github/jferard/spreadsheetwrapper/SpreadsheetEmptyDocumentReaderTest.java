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

import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.NoSuchElementException;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

@SuppressWarnings("unused")
public abstract class SpreadsheetEmptyDocumentReaderTest {
	protected SpreadsheetDocumentFactory factory;
	protected SpreadsheetDocumentReader sdr;

	@Before
	public abstract void setUp();

	@After
	public abstract void tearDown();

	@Test(expected = IndexOutOfBoundsException.class)
	public final void testSheet() {
		SpreadsheetReader sr = this.sdr.getSpreadsheet(0);
	}

	@Test(expected = NoSuchElementException.class)
	public final void testSheet3() {
		final SpreadsheetReader sr = this.sdr.getSpreadsheet("Feuille1");
	}

	@Test(expected = IndexOutOfBoundsException.class)
	public final void testSheet5() throws SpreadsheetException {
		final SpreadsheetReader sr = this.sdr.getSpreadsheet(1);
	}

	@Test(expected = IndexOutOfBoundsException.class)
	public final void testSheet6() throws SpreadsheetException {
		final SpreadsheetReader sr = this.sdr.getSpreadsheet(-1);
	}

	@Test(expected = NoSuchElementException.class)
	public final void testSheet7() throws SpreadsheetException {
		final SpreadsheetReader sr = this.sdr.getSpreadsheet("F");
	}

	@Test(expected = IndexOutOfBoundsException.class)
	public final void testSheetCursors() throws SpreadsheetException {
		final SpreadsheetReaderCursor c1 = this.sdr.getNewCursorByIndex(0);
	}

	@Test(expected = NoSuchElementException.class)
	public final void testSheetCursors2() throws SpreadsheetException {
		final SpreadsheetReaderCursor c2 = this.sdr
				.getNewCursorByName("Feuille1");
	}

	@Test(expected = NoSuchElementException.class)
	public final void testSheetDoesntExist() throws SpreadsheetException {
		final SpreadsheetReader sr = this.sdr.getSpreadsheet("test");
	}

	@Test
	public final void testSheetNames() {
		final List<String> names = this.sdr.getSheetNames();
		Assert.assertEquals(Collections.emptyList(), names);
	}

	@Test(expected = IndexOutOfBoundsException.class)
	public final void testSheetOutOfBounds2() throws SpreadsheetException {
		final SpreadsheetReader sr = this.sdr.getSpreadsheet(-1);
	}

	protected abstract String getExtension();

	protected abstract SpreadsheetDocumentFactory getFactory();
}