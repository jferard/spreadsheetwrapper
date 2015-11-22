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

import com.github.jferard.spreadsheetwrapper.impl.StyleUtility;

/**
 * @author Julien
 *
 */
public abstract class AbstractSpreadsheetEmptyWriterTest {
	private static final String STYLE_NAME = "mystyle2";

	/** name of the test */
	@Rule
	public TestName name = new TestName();

	/** logger, static initialization */
	private final Logger logger = Logger.getLogger(this.getClass().getName());

	/** factory */
	protected SpreadsheetDocumentFactory factory;

	/** document writer */
	protected SpreadsheetDocumentWriter sdw;

	/** the writer */
	protected SpreadsheetWriter sw;

	/** set the test up */
	@Before
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			this.sdw = this.factory.create();
			this.sw = this.sdw.addSheet(0, "first sheet");
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.INFO, "", e);
			Assert.fail();
		}
	}

	/** tear the test down */
	@After
	public void tearDown() {
		try {
			final File outputFile = SpreadsheetTestHelper.getOutputFile(this
					.getClass().getSimpleName(), this.name.getMethodName(),
					this.getProperties().getExtension());
			this.sdw.saveAs(outputFile);
			this.sdw.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.INFO, "", e);
			Assert.fail();
		}
	}

	/**
	 * must create a A1 blue bold head cell and a C3 simple tail cell
	 */
	@Test
	public final void testCreate2StylesFromObjectsAndSetStylesWithGap() {
		final WrapperFont wrapperFont = new WrapperFont().setBold();
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle(
				WrapperColor.AQUA, wrapperFont);
		this.sdw.setStyle(AbstractSpreadsheetEmptyWriterTest.STYLE_NAME,
				wrapperCellStyle);
		this.sw.setCellContent(0, 0, "Head",
				AbstractSpreadsheetEmptyWriterTest.STYLE_NAME);
		this.sw.setCellContent(2, 2, "Tail");

		try {
			for (int r = 0; r <= 3; r++) {
				for (int c = 0; c <= 3; c++) {
					final WrapperCellStyle otherCellStyle = this.sw.getStyle(r,
							c);
					WrapperCellStyle expectedCellStyle = null;
					if (r == 0 && c == 0)
						expectedCellStyle = wrapperCellStyle;
					else if (r == 2 && c <= 2)
						expectedCellStyle = WrapperCellStyle.EMPTY;
					Assert.assertEquals(String.format("%d * %d", r, c),
							expectedCellStyle, otherCellStyle);
				}
			}
		} catch (final UnsupportedOperationException e) {
			this.logger.log(Level.INFO, "", e);
			Assume.assumeNoException(e);
		}
	}

	/**
	 * must create a A1 blue bold head cell and a A2 simple tail cell
	 */
	@Test
	public final void testCreate2StylesFromObjectsAndSetStylesWithNoGap() { // buggy
		// for
		// simpleodf
		final WrapperFont wrapperFont = new WrapperFont().setBold();
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle(
				WrapperColor.AQUA, wrapperFont);
		this.sdw.setStyle(AbstractSpreadsheetEmptyWriterTest.STYLE_NAME,
				wrapperCellStyle);
		this.sw.setCellContent(0, 0, "Head",
				AbstractSpreadsheetEmptyWriterTest.STYLE_NAME);
		this.sw.setCellContent(1, 0, "Tail");

		try {
			final WrapperCellStyle newCellStyle = this.sw.getStyle(0, 0);
			Assert.assertEquals(wrapperCellStyle, newCellStyle);
		} catch (final UnsupportedOperationException e) {
			this.logger.log(Level.INFO, "", e);
			Assume.assumeNoException(e);
		}
	}

	/**
	 * must create a A1 blue bold head cell and a A2 simple tail cell
	 */
	@Test
	public final void testCreateStyleFromObjectsAndSetStyle() {
		final WrapperFont wrapperFont = new WrapperFont().setBold();
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle(
				WrapperColor.AQUA, wrapperFont);
		this.sdw.setStyle(AbstractSpreadsheetEmptyWriterTest.STYLE_NAME,
				wrapperCellStyle);
		this.sw.setCellContent(0, 0, "Head",
				AbstractSpreadsheetEmptyWriterTest.STYLE_NAME);
		this.sw.setCellContent(1, 0, "Tail", "");

		try {
			final WrapperCellStyle newCellStyle = this.sw.getStyle(0, 0);
			Assert.assertEquals(wrapperCellStyle, newCellStyle);
		} catch (final UnsupportedOperationException e) {
			this.logger.log(Level.INFO, "", e);
			Assume.assumeNoException(e);
		}
	}

	/**
	 * must create a A1 blue bold head cell
	 */
	@Test
	public final void testCreateStyleFromStringAndSetStyle() {
		final StyleUtility utility = new StyleUtility();
		final WrapperCellStyle wrapperCellStyle = utility
				.toWrapperCellStyle("background-color:#999999;font-weight:bold");
		this.sdw.setStyle(AbstractSpreadsheetEmptyWriterTest.STYLE_NAME,
				wrapperCellStyle);
		this.sw.setCellContent(0, 0, "Head",
				AbstractSpreadsheetEmptyWriterTest.STYLE_NAME);

		try {
			final WrapperCellStyle newCellStyle = this.sw.getStyle(0, 0);
			Assert.assertEquals(wrapperCellStyle, newCellStyle);
		} catch (final UnsupportedOperationException e) {
			this.logger.log(Level.INFO, "", e);
			Assume.assumeNoException(e);
		}
	}

	/**
	 * set a boolean on A1 and test if it is here
	 */
	@Test
	public void testSetBoolean() {
		try {
			this.sw.setCellContent(0, 0, true);
			Assert.assertEquals(true, this.sw.getBoolean(0, 0));

			this.sw.setBoolean(0, 0, true);
			Assert.assertEquals(true, this.sw.getCellContent(0, 0));
		} catch (final UnsupportedOperationException e) {
			this.logger.log(Level.INFO, "", e);
			Assume.assumeNoException(e);
		}
	}

	/**
	 * set a date on A1 and test if it is here the date can be truncated to day
	 * or second
	 */
	@Test
	public void testSetDate() {
		final long DAY = 52 * 86400000L; // 52 days = 4492800000 ms
		final long HOUR_MIN_AND_SEC = ((3 * 60 + 25) * 60 + 45) * 1000L; // 3 h
																			// 25
																			// min
																			// 45
																			// s
																			// =
																			// 12345000
																			// ms
		final long MILLIS = 789L; // 789 ms

		final Date dateTruncatedToDay = new Date(DAY);
		final Date dateTruncatedtoSecond = new Date(DAY + HOUR_MIN_AND_SEC);
		final Date dateNotTruncated = new Date(DAY + HOUR_MIN_AND_SEC + MILLIS);

		final Date d2 = this.sw.setDate(2, 2, dateNotTruncated);
		Assert.assertTrue(String.format("%d not in (%d, %d, %d)", d2.getTime(),
				dateNotTruncated.getTime(), dateTruncatedtoSecond.getTime(),
				dateTruncatedToDay.getTime()),
						d2.equals(dateNotTruncated) || d2.equals(dateTruncatedtoSecond)
						|| d2.equals(dateTruncatedToDay));
		Assert.assertEquals(d2, this.sw.getDate(2, 2));
		Assert.assertEquals(d2, this.sw.getCellContent(2, 2));

		final Object o2 = this.sw.setCellContent(2, 2, dateNotTruncated);
		Assert.assertTrue(o2.equals(dateNotTruncated)
				|| o2.equals(dateTruncatedtoSecond)
				|| o2.equals(dateTruncatedToDay));
		Assert.assertEquals(o2, this.sw.getDate(2, 2));
		Assert.assertEquals(o2, this.sw.getCellContent(2, 2));
	}

	/**
	 * set a double on C3 and test if it is here
	 */
	@Test
	public void testSetDouble() {
		this.sw.setCellContent(2, 2, 100.7);
		Assert.assertEquals(Double.valueOf(100.7), this.sw.getDouble(2, 2));

		this.sw.setDouble(2, 2, 100.7);
		Assert.assertEquals(Double.valueOf(100.7), this.sw.getCellContent(2, 2));
	}

	/**
	 * set a formula on B2 and test if it is here
	 */
	@Test
	public void testSetFormula() {
		try {
			this.sw.setFormula(1, 1, "A1");
			Assert.assertEquals("A1", this.sw.getFormula(1, 1));

			this.sw.setFormula(1, 1, "A1");
			Assert.assertEquals("A1", this.sw.getCellContent(1, 1));
		} catch (final UnsupportedOperationException e) {
			this.logger.log(Level.INFO, "", e);
			Assume.assumeNoException(e);
		}
	}

	/**
	 * set an incorrect formula on B2 and test if it is here
	 */
	@Test
	public void testSetIncorrrectFormula() {
		try {
			final String formula = "afdferg'|[{*dfgdrsg]";
			this.sw.setFormula(1, 1, formula);
			Assert.assertEquals(formula, this.sw.getFormula(1, 1));

			this.sw.setFormula(1, 1, formula);
			Assert.assertEquals(formula, this.sw.getCellContent(1, 1));
		} catch (final IllegalArgumentException e) {
			Assume.assumeNoException(e);
		} catch (final UnsupportedOperationException e) {
			this.logger.log(Level.INFO, "", e);
			Assume.assumeNoException(e);
		}
	}

	/**
	 * set an integer on B2 and test if it is here
	 */
	@Test
	public void testSetInteger() {
		this.sw.setCellContent(1, 1, 10);
		Assert.assertEquals(Integer.valueOf(10), this.sw.getInteger(1, 1));

		this.sw.setInteger(1, 1, 10);
		Assert.assertEquals(Integer.valueOf(10), this.sw.getCellContent(1, 1));
	}

	/**
	 * set a text on B2 and test if it is here
	 */
	@Test
	public void testSetText() {
		this.sw.setCellContent(1, 1, "10");
		Assert.assertEquals("10", this.sw.getText(1, 1));

		this.sw.setText(1, 1, "10");
		Assert.assertEquals("10", this.sw.getCellContent(1, 1));
	}

	/**
	 * set a text on B2 to B1001 and test if it is here
	 */
	@Test
	public void testSetTextOnCol0To1000() {
		int i = 0;
		try {
			for (i = 0; i < 1000; i += 100) {
				this.sw.setText(1, i, "10");
				Assert.assertEquals("10", this.sw.getText(1, i));
			}
		} catch (final UnsupportedOperationException e) {
			this.logger.log(Level.INFO, "", e);
			if (i >= 255)
				Assume.assumeNoException(e);
			else
				throw e;
		}
	}

	/**
	 * set a text on B256 and test if it is here
	 */
	@Test
	public void testSetTextOnCol255() {
		this.sw.setText(1, 255, "100");
		Assert.assertEquals("100", this.sw.getText(1, 255));
	}

	/**
	 * set a text on B(-9) and test if it is here. Must throw an exception
	 */
	@Test(expected = IllegalArgumentException.class)
	public void testSetTextOnColMinus10() {
		this.sw.setText(1, -10, "10");
		Assert.assertEquals("10", this.sw.getText(1, -10));
	}

	/**
	 * set a text on B100000 and test if it is here
	 */
	@Test
	public void testSetTextOnRow0To100000() {
		int j = 0;
		try {
			for (j = 0; j < 100000; j += 20000) {
				this.sw.setText(j, 1, "10");
				Assert.assertEquals("10", this.sw.getText(j, 1));
			}
		} catch (final UnsupportedOperationException e) {
			this.logger.log(Level.INFO, "", e);
			if (j >= 65535)
				Assume.assumeNoException(e);
			else
				throw e;
		}
	}

	/**
	 * set a text on B51 and test if it is here
	 */
	@Test
	public void testSetTextOnRow50() {
		this.sw.setText(50, 1, "10");
		Assert.assertEquals("10", this.sw.getText(50, 1));
	}

	/**
	 * set a text on B65536 and test if it is here
	 */
	@Test
	public void testSetTextOnRow65535() {
		this.sw.setText(0, 1, "0");
		this.sw.setText(65535, 1, "100");
		Assert.assertEquals("100", this.sw.getText(65535, 1));
	}

	/**
	 * set a text on row -10 and test if it is here
	 */
	@Test(expected = IllegalArgumentException.class)
	public void testSetTextOnRowMinus10() {
		this.sw.setText(-10, 1, "10");
		Assert.assertEquals("10", this.sw.getText(-10, 1));
	}

	/**
	 * @return the properties of the API for the test
	 */
	protected abstract TestProperties getProperties();
}
