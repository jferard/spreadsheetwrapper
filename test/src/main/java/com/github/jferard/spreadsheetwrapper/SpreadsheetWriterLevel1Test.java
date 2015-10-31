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
import java.util.Date;

import org.junit.After;
import org.junit.Assert;
import org.junit.Assume;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;

public abstract class SpreadsheetWriterLevel1Test extends SpreadsheetReaderTest {
	@Rule
	public TestName name = new TestName();
	protected SpreadsheetDocumentWriter sdw;

	protected SpreadsheetWriter sw;

	/** set the test up */
	@Before
	@Override
	@SuppressWarnings("nullness")
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			final URL resourceURL = this.getProperties().getSourceURL();
			Assume.assumeNotNull(resourceURL);

			final InputStream inputStream = resourceURL.openStream();
			this.sdw = this.factory.openForWrite(inputStream);
			this.sdr = this.sdw;
			this.sw = this.sdw.getSpreadsheet(0);
			this.sr = this.sw;
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		} catch (final IOException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	/** tear the test down */
	@After
	@Override
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

	@Test(expected = IllegalArgumentException.class)
	public final void testCellDateToInteger() {
		this.sw.setDate(0, 1, new Date(0)); // setDate : 0 UTC = 1 CET
		Assert.assertEquals((Integer) 0, this.sw.getInteger(0, 1));
		Assert.fail();
	}

	@Test
	public final void testCellDateZero() {
		try {
			final Date d = this.sw.setDate(0, 1, new Date(0)); // setDate : 0
			// UTC = 1 CET
			Assert.assertEquals(d, new Date(0));
			Assert.assertEquals(d, this.sw.getDate(0, 1)); // getDate
			// : 0 CET = 0 CET
		} catch (final IllegalArgumentException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		}
	}

	@Test
	public void testSetBoolean() {
		final int r = 5;
		final int c = 6;
		try {
			this.sw.setBoolean(r, c, true);
			Assert.assertEquals(true, this.sw.getBoolean(r, c));
		} catch (final IllegalArgumentException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		}
	}

	@Test
	public final void testSetCellDate() {
		final int r = 5;
		final int c = 6;
		try {
			final Date dm = new Date(1234567891);
			final Date ds = new Date(1234567891 - Math.floorMod(1234567891,
					1000));
			final Date dd = new Date(1234567891 - Math.floorMod(1234567891,
					1000 * 86400));

			final Date d2 = this.sw.setDate(r, c, dm); // setDate :
			// 0
			// UTC = 1 CET
			Assert.assertTrue(d2.equals(dm) || d2.equals(ds) || d2.equals(dd));
			final Date d3 = this.sw.getDate(r, c);
			Assert.assertEquals(d2, d3);
		} catch (final IllegalArgumentException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		}
	}
}
