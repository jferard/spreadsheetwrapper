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
import java.util.logging.Level;
import java.util.logging.Logger;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;

/**
 * The AbstractSpreadSheetEmptyWriterCountsTest class tests the row / cell
 * counts computation by a wrapper. Those tests are now successful for all
 * wrappers.
 */
public abstract class AbstractSpreadSheetEmptyWriterCountsTest {
	/** name of the test */
	@Rule
	public TestName name = new TestName();

	/** logger, static initialization */
	private final Logger logger = Logger.getLogger(this.getClass().getName());

	/** the document writer */
	protected SpreadsheetDocumentWriter documentWriter;

	/** the factory */
	protected SpreadsheetDocumentFactory factory;

	/** the sheet writer */
	protected SpreadsheetWriter sheetWriter;

	private static final int BIG_INDEX = 9999;

	/** set the test up */
	@Before
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			this.documentWriter = this.factory.create();
			this.sheetWriter = this.documentWriter.addSheet(0, "first sheet");
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** tear the test down */
	@After
	public void tearDown() {
		try {
			final File outputFile = SpreadsheetTestHelper.getOutputFile(
					this.factory, this.getClass().getSimpleName(),
					this.name.getMethodName());
			this.documentWriter.saveAs(outputFile);
			this.documentWriter.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** a get doesn't change the row count */
	@Test
	public void testGetIsNotCreateGetThenNothing() {
		this.sheetWriter.getCellContent(15, 5);
		Assert.assertEquals(0, this.sheetWriter.getRowCount());
	}

	/** a get doesn't change the row count, but a set does. Get before set. */
	@Test
	public void testGetIsNotCreateGetThenSet() {
		this.sheetWriter.getCellContent(15, 5);
		this.sheetWriter.setText(1, 3, "a");
		Assert.assertEquals(2, this.sheetWriter.getRowCount());
		Assert.assertEquals(4, this.sheetWriter.getCellCount(1));
	}

	/** a get doesn't change the row count, but a set does. Set before get. */
	@Test
	public void testGetIsNotCreateSetThenGet() {
		this.sheetWriter.setText(1, 3, "a");
		this.sheetWriter.getCellContent(15, 5);
		Assert.assertEquals(2, this.sheetWriter.getRowCount());
		Assert.assertEquals(4, this.sheetWriter.getCellCount(1));
	}

	/** */
	@Test
	public void testGetObjetAtBigIndex() {
		this.sheetWriter.getCellContent(0, BIG_INDEX);
	}

	/** */
	@Test(expected = IllegalArgumentException.class)
	public void testGetObjetAtNegativeIndex() {
		this.sheetWriter.getCellContent(0, -1);
	}

	/** a set does not change the col count of the next row */
	@Test
	public void testNonExistingRowCellCountK11() {
		this.sheetWriter.setText(10, 10, "10:10");
		Assert.assertEquals(0, this.sheetWriter.getCellCount(11));
	}

	/**
	 * set on K11 -> row count = 11, col count from row 11 = 11, other col count
	 * = 0
	 */
	@Test
	public void testRowAndCellCountsK11() {
		this.sheetWriter.setText(10, 10, "10:10");
		Assert.assertEquals(11, this.sheetWriter.getRowCount());
		Assert.assertEquals(0, this.sheetWriter.getCellCount(1));
		Assert.assertEquals(0, this.sheetWriter.getCellCount(9));
		Assert.assertEquals(11, this.sheetWriter.getCellCount(10));
	}

	/** set on A1 -> row count = 1, col count from row 1 = 1 */
	@Test
	public void testRowCountA1() {
		this.sheetWriter.setInteger(0, 0, 1);
		Assert.assertEquals(1, this.sheetWriter.getRowCount());
		Assert.assertEquals(1, this.sheetWriter.getCellCount(0));
	}

	/** set on D2 -> row count = 2, col count from row 2 = 4 */
	@Test
	public void testRowCountD2() {
		this.sheetWriter.setInteger(1, 3, 1);
		Assert.assertEquals(2, this.sheetWriter.getRowCount());
		Assert.assertEquals(4, this.sheetWriter.getCellCount(1));
	}

	/** */
	@Test
	public void testSetObjetAtBigIndex() {
		Assert.assertEquals(Integer.valueOf(1), this.sheetWriter.setCellContent(0, BIG_INDEX,
				Integer.valueOf(1)));
	}

	/** */
	@Test(expected = IllegalArgumentException.class)
	public void testSetObjetAtNegativeIndex() {
		this.sheetWriter.setCellContent(0, -1, Integer.valueOf(1));
	}

	/**
	 * a beginning, row count = 0
	 *
	 * @Test public void testRowCountZero() { Assert.assertEquals(0,
	 *       this.sheetWriter.getRowCount()); }
	 *
	 *       /** get the properties
	 */
	protected abstract TestProperties getProperties();

}
