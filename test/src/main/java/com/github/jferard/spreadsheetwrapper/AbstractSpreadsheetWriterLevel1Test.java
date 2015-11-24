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
import java.util.logging.Level;
import java.util.logging.Logger;

import org.junit.After;
import org.junit.Assert;
import org.junit.Assume;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;

public abstract class AbstractSpreadsheetWriterLevel1Test extends
AbstractSpreadsheetReaderTest {
	/** name of the test */
	@Rule
	public TestName name = new TestName();
	/** logger, static initialization */
	private final Logger logger = Logger.getLogger(this.getClass().getName());

	/** the document writer */
	protected SpreadsheetDocumentWriter documentWriter;

	/** the sheet writer */
	protected SpreadsheetWriter sheetWriter;

	/** set the test up */
	@Before
	@Override
	@SuppressWarnings("nullness")
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			final URL resourceURL = this.getProperties().getResourceURL();
			Assume.assumeNotNull(resourceURL);

			final InputStream inputStream = resourceURL.openStream();
			this.documentWriter = this.factory.openForWrite(inputStream);
			this.documentReader = this.documentWriter;
			this.sheetWriter = this.documentWriter.getSpreadsheet(0);
			this.sheetReader = this.sheetWriter;
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		} catch (final IOException e) {
			this.logger.log(Level.WARNING, "", e);
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
			this.documentWriter.saveAs(outputFile);
			this.documentWriter.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** can't convert date to integer */
	@Test(expected = IllegalArgumentException.class)
	public final void testCellDateToInteger() {
		this.sheetWriter.setDate(0, 1, new Date(0)); // setDate : 0 UTC = 1 CET
		Assert.assertEquals((Integer) 0, this.sheetWriter.getInteger(0, 1));
		Assert.fail();
	}

	/** set and get date on cell A2 */
	@Test
	public final void testCellDateZero() {
		try {
			final Date date = this.sheetWriter.setDate(0, 1, new Date(0)); // setDate
																			// :
																			// 0
			// UTC = 1 CET
			Assert.assertEquals(date, new Date(0));
			Assert.assertEquals(date, this.sheetWriter.getDate(0, 1)); // getDate
			// : 0 CET = 0 CET
		} catch (final IllegalArgumentException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		}
	}

	/** set and get boolean on cell G6 */
	@Test
	public void testSetBoolean() {
		final int r = 5;
		final int c = 6;
		try {
			this.sheetWriter.setBoolean(r, c, true);
			Assert.assertEquals(true, this.sheetWriter.getBoolean(r, c));
		} catch (final IllegalArgumentException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		}
	}

	/** set and get date G6 */
	@Test
	public final void testSetCellDate() {
		final int r = 5;
		final int c = 6;
		try {
			final Date dateMillis = new Date(1234567891);
			final Date dateSeconds = new Date(1234567891 - Math.floorMod(
					1234567891, 1000));
			final Date dateDays = new Date(1234567891 - Math.floorMod(
					1234567891, 1000 * 86400));

			final Date dateSet = this.sheetWriter.setDate(r, c, dateMillis); // setDate
																				// :
			// 0
			// UTC = 1 CET
			Assert.assertTrue(dateSet.equals(dateMillis)
					|| dateSet.equals(dateSeconds) || dateSet.equals(dateDays));
			final Date dateRead = this.sheetWriter.getDate(r, c);
			Assert.assertEquals(dateSet, dateRead);
		} catch (final IllegalArgumentException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		}
	}
}
