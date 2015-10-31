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
public abstract class SpreadsheetDocumentReaderTest {
	protected SpreadsheetDocumentFactory factory;
	protected SpreadsheetDocumentReader sdr;

	@Before
	@SuppressWarnings("nullness")
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			final URL sourceURL = this.getProperties().getSourceURL();
			Assume.assumeNotNull(sourceURL);

			final InputStream inputStream = sourceURL.openStream();
			this.sdr = this.factory.openForRead(inputStream);
			Assert.assertEquals(1, this.sdr.getSheetCount());
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		} catch (final IOException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@After
	public void tearDown() {
		try {
			this.sdr.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test
	public final void testGetSheetNames() {
		final List<String> names = this.sdr.getSheetNames();
		Assert.assertEquals(Arrays.asList("Feuille1"), names);
	}

	@Test(expected = IndexOutOfBoundsException.class)
	public final void testGetSheetNumberMinusOneThrowsIndexOutOfBounds()
			throws SpreadsheetException {
		final SpreadsheetReader sr = this.sdr.getSpreadsheet(-1);
	}

	@Test(expected = IndexOutOfBoundsException.class)
	public final void testGetSheetNumberOneThrowsIndexOutOfBounds()
			throws SpreadsheetException {
		final SpreadsheetReader sr = this.sdr.getSpreadsheet(1);
	}

	@Test
	public final void testGetSheetTwice() throws SpreadsheetException {
		SpreadsheetReader sr = this.sdr.getSpreadsheet("Feuille1");
		sr = this.sdr.getSpreadsheet("Feuille1");
	}

	@Test
	public final void testSheetByIndex() {
		try {
			SpreadsheetReader sr = this.sdr.getSpreadsheet(0);
			sr = this.sdr.getSpreadsheet(0);
			Assert.assertEquals("Feuille1", sr.getName());
			Assert.assertEquals(117, sr.getRowCount());
			Assert.assertEquals(7, sr.getCellCount(0));
		} catch (final NoSuchElementException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test
	public final void testSheetByIndexAndName() {
		try {
			final SpreadsheetReader sr1 = this.sdr.getSpreadsheet("Feuille1");
			final SpreadsheetReader sr2 = this.sdr.getSpreadsheet(0); // use
			// Accessor
			Assert.assertEquals(sr1, sr2);
		} catch (final NoSuchElementException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test
	public final void testSheetByName() {
		try {
			final SpreadsheetReader sr = this.sdr.getSpreadsheet("Feuille1");
			Assert.assertEquals("Feuille1", sr.getName());
			Assert.assertEquals(117, sr.getRowCount());
			Assert.assertEquals(7, sr.getCellCount(0));
		} catch (final NoSuchElementException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test(expected = NoSuchElementException.class)
	public final void testSheetDoesntExist() throws SpreadsheetException {
		final SpreadsheetReader sr = this.sdr.getSpreadsheet("test");
	}

	@Test
	public final void testSheetGetCursors() {
		try {
			final SpreadsheetReaderCursor c1 = this.sdr.getNewCursorByIndex(0);
			final SpreadsheetReaderCursor c2 = this.sdr
					.getNewCursorByName("Feuille1");
		} catch (final SpreadsheetException e) {
			Assert.fail();
		}
	}

	protected abstract TestProperties getProperties();
}