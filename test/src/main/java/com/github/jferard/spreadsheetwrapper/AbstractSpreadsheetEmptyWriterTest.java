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

import com.github.jferard.spreadsheetwrapper.xls.XlsConstants;

/**
 * @author Julien
 *
 */
public abstract class AbstractSpreadsheetEmptyWriterTest {
	/** the style name */
	private static final String STYLE_NAME = "mystyle2";

	/** name of the test */
	@Rule
	public TestName name = new TestName();

	/** logger, static initialization */
	private final Logger logger = Logger.getLogger(this.getClass().getName());

	/** document writer */
	protected SpreadsheetDocumentWriter documentWriter;

	/** factory */
	protected SpreadsheetDocumentFactory factory;

	/** the writer */
	protected SpreadsheetWriter sheetWriter;

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
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle()
		.setBackgroundColor(WrapperColor.AQUA).setCellFont(wrapperFont);
		this.documentWriter
		.setStyle(AbstractSpreadsheetEmptyWriterTest.STYLE_NAME,
				wrapperCellStyle);
		this.sheetWriter.setCellContent(0, 0, "Head",
				AbstractSpreadsheetEmptyWriterTest.STYLE_NAME);
		this.sheetWriter.setCellContent(2, 2, "Tail");

		try {
			for (int r = 0; r <= 3; r++) {
				for (int c = 0; c <= 3; c++) {
					final WrapperCellStyle otherCellStyle = this.sheetWriter
							.getStyle(r, c);
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
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle()
		.setBackgroundColor(WrapperColor.AQUA).setCellFont(wrapperFont);
		this.documentWriter
		.setStyle(AbstractSpreadsheetEmptyWriterTest.STYLE_NAME,
				wrapperCellStyle);
		this.sheetWriter.setCellContent(0, 0, "Head",
				AbstractSpreadsheetEmptyWriterTest.STYLE_NAME);
		this.sheetWriter.setCellContent(1, 0, "Tail");

		try {
			final WrapperCellStyle newCellStyle = this.sheetWriter.getStyle(0,
					0);
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
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle()
		.setBackgroundColor(WrapperColor.AQUA).setCellFont(wrapperFont);
		this.documentWriter
		.setStyle(AbstractSpreadsheetEmptyWriterTest.STYLE_NAME,
				wrapperCellStyle);
		this.sheetWriter.setCellContent(0, 0, "Head",
				AbstractSpreadsheetEmptyWriterTest.STYLE_NAME);
		this.sheetWriter.setCellContent(1, 0, "Tail", "");

		try {
			final WrapperCellStyle newCellStyle = this.sheetWriter.getStyle(0,
					0);
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
	public final void testCreateStyleFromObjectsAndSetStyleFontFamily() {
		final WrapperFont wrapperFont = new WrapperFont().setItalic()
				.setSize(20).setFamily(WrapperFont.COURIER_NAME);
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle()
		.setBorderLineWidth(WrapperCellStyle.THIN_LINE).setCellFont(
						wrapperFont);
		this.documentWriter
		.setStyle(AbstractSpreadsheetEmptyWriterTest.STYLE_NAME,
				wrapperCellStyle);
		try {
			this.sheetWriter.setCellContent(0, 0, "From name",
					AbstractSpreadsheetEmptyWriterTest.STYLE_NAME);
			final WrapperCellStyle newCellStyle = this.sheetWriter.getStyle(0,
					0);
			Assert.assertEquals(wrapperCellStyle, newCellStyle);
		} catch (final UnsupportedOperationException e) {
			this.logger.log(Level.INFO, "", e);
			Assume.assumeNoException(e);
		}

		try {
			this.sheetWriter.setCellContent(1, 0, "From Style");
			this.sheetWriter.setStyle(1, 0, wrapperCellStyle);
			final WrapperCellStyle newCellStyle = this.sheetWriter.getStyle(1,
					0);
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
	public final void testCreateStyleFromObjectsAndSetStyleItalics() {
		final WrapperFont wrapperFont = new WrapperFont().setItalic()
				.setSize(20).setColor(WrapperColor.ORANGE);
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle()
		.setCellFont(wrapperFont);
		this.documentWriter
		.setStyle(AbstractSpreadsheetEmptyWriterTest.STYLE_NAME,
				wrapperCellStyle);
		try {
			this.sheetWriter.setCellContent(0, 0, "From name",
					AbstractSpreadsheetEmptyWriterTest.STYLE_NAME);
			final WrapperCellStyle newCellStyle = this.sheetWriter.getStyle(0,
					0);
			Assert.assertEquals(wrapperCellStyle, newCellStyle);
		} catch (final UnsupportedOperationException e) {
			this.logger.log(Level.INFO, "", e);
			Assume.assumeNoException(e);
		}

		try {
			this.sheetWriter.setCellContent(1, 0, "From Style");
			this.sheetWriter.setStyle(1, 0, wrapperCellStyle);
			final WrapperCellStyle newCellStyle = this.sheetWriter.getStyle(1,
					0);
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
		this.documentWriter
		.setStyle(AbstractSpreadsheetEmptyWriterTest.STYLE_NAME,
				wrapperCellStyle);
		this.sheetWriter.setCellContent(0, 0, "Head",
				AbstractSpreadsheetEmptyWriterTest.STYLE_NAME);

		try {
			final WrapperCellStyle newCellStyle = this.sheetWriter.getStyle(0,
					0);
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
			this.sheetWriter.setCellContent(0, 0, true);
			Assert.assertEquals(true, this.sheetWriter.getBoolean(0, 0));

			this.sheetWriter.setBoolean(0, 0, true);
			Assert.assertEquals(true, this.sheetWriter.getCellContent(0, 0));
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

		final Date dateSet = this.sheetWriter.setDate(2, 2, dateNotTruncated);
		Assert.assertTrue(
				String.format("%d not in (%d, %d, %d)", dateSet.getTime(),
						dateNotTruncated.getTime(),
						dateTruncatedtoSecond.getTime(),
						dateTruncatedToDay.getTime()),
						dateSet.equals(dateNotTruncated)
						|| dateSet.equals(dateTruncatedtoSecond)
						|| dateSet.equals(dateTruncatedToDay));
		Assert.assertEquals(dateSet, this.sheetWriter.getDate(2, 2));
		Assert.assertEquals(dateSet, this.sheetWriter.getCellContent(2, 2));

		final Object dateSetAsObject = this.sheetWriter.setCellContent(2, 2,
				dateNotTruncated);
		Assert.assertTrue(dateSetAsObject.equals(dateNotTruncated)
				|| dateSetAsObject.equals(dateTruncatedtoSecond)
				|| dateSetAsObject.equals(dateTruncatedToDay));
		Assert.assertEquals(dateSetAsObject, this.sheetWriter.getDate(2, 2));
		Assert.assertEquals(dateSetAsObject,
				this.sheetWriter.getCellContent(2, 2));
	}

	/**
	 * set a double on C3 and test if it is here
	 */
	@Test
	public void testSetDouble() {
		this.sheetWriter.setCellContent(2, 2, 100.7);
		Assert.assertEquals(Double.valueOf(100.7),
				this.sheetWriter.getDouble(2, 2));

		this.sheetWriter.setDouble(2, 2, 100.7);
		Assert.assertEquals(Double.valueOf(100.7),
				this.sheetWriter.getCellContent(2, 2));
	}

	/**
	 * set a formula on B2 and test if it is here
	 */
	@Test
	public void testSetFormula() {
		try {
			this.sheetWriter.setFormula(1, 1, "A1");
			Assert.assertEquals("A1", this.sheetWriter.getFormula(1, 1));

			this.sheetWriter.setFormula(1, 1, "A1");
			Assert.assertEquals("A1", this.sheetWriter.getCellContent(1, 1));
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
			this.sheetWriter.setFormula(1, 1, formula);
			Assert.assertEquals(formula, this.sheetWriter.getFormula(1, 1));

			this.sheetWriter.setFormula(1, 1, formula);
			Assert.assertEquals(formula, this.sheetWriter.getCellContent(1, 1));
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
		this.sheetWriter.setCellContent(1, 1, 10);
		Assert.assertEquals(Integer.valueOf(10),
				this.sheetWriter.getInteger(1, 1));

		this.sheetWriter.setInteger(1, 1, 10);
		Assert.assertEquals(Integer.valueOf(10),
				this.sheetWriter.getCellContent(1, 1));
	}

	/**
	 * set a text on B2 and test if it is here
	 */
	@Test
	public void testSetText() {
		this.sheetWriter.setCellContent(1, 1, "10");
		Assert.assertEquals("10", this.sheetWriter.getText(1, 1));

		this.sheetWriter.setText(1, 1, "10");
		Assert.assertEquals("10", this.sheetWriter.getCellContent(1, 1));
	}

	/**
	 * set a text on B2 to B1001 and test if it is here
	 */
	@Test
	public void testSetTextOnCol0To1000() {
		int colIndex = 0;
		try {
			for (colIndex = 0; colIndex < 1000; colIndex += 100) {
				this.sheetWriter.setText(1, colIndex, "10");
				Assert.assertEquals("10", this.sheetWriter.getText(1, colIndex));
			}
		} catch (final UnsupportedOperationException e) {
			this.logger.log(Level.INFO, "", e);
			if (colIndex >= XlsConstants.MAX_COLUMNS)
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
		final String HUNDRED = "100";
		final String text = this.sheetWriter.setText(1,
				XlsConstants.MAX_COLUMNS - 1, HUNDRED);
		Assert.assertEquals(HUNDRED, text);
		Assert.assertEquals(HUNDRED,
				this.sheetWriter.getText(1, XlsConstants.MAX_COLUMNS - 1));
	}

	/**
	 * set a text on B(-9) and test if it is here. Must throw an exception
	 */
	@Test(expected = IllegalArgumentException.class)
	public void testSetTextOnColMinus10() {
		this.sheetWriter.setText(1, -10, "10");
		Assert.assertEquals("10", this.sheetWriter.getText(1, -10));
	}

	/**
	 * set a text on B100000 and test if it is here
	 */
	@Test
	public void testSetTextOnRow0To100000() {
		int rowIndex = 0;
		try {
			for (rowIndex = 0; rowIndex < 100000; rowIndex += 20000) {
				this.sheetWriter.setText(rowIndex, 1, "10");
				Assert.assertEquals("10", this.sheetWriter.getText(rowIndex, 1));
			}
		} catch (final UnsupportedOperationException e) {
			this.logger.log(Level.INFO, "", e);
			if (rowIndex >= XlsConstants.MAX_ROWS_PER_SHEET)
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
		this.sheetWriter.setText(50, 1, "10");
		Assert.assertEquals("10", this.sheetWriter.getText(50, 1));
	}

	/**
	 * set a text on B65536 and test if it is here
	 */
	@Test
	public void testSetTextOnRow65535() {
		this.sheetWriter.setText(0, 1, "0");
		final String setAndExpected = "100";
		final String set = this.sheetWriter.setText(
				XlsConstants.MAX_ROWS_PER_SHEET - 1, 1, setAndExpected);
		Assert.assertEquals(setAndExpected, set);
		Assert.assertEquals(setAndExpected, this.sheetWriter.getText(
				XlsConstants.MAX_ROWS_PER_SHEET - 1, 1));
	}

	/**
	 * set a text on row -10 and test if it is here
	 */
	@Test(expected = IllegalArgumentException.class)
	public void testSetTextOnRowMinus10() {
		this.sheetWriter.setText(-10, 1, "10");
		Assert.assertEquals("10", this.sheetWriter.getText(-10, 1));
	}

	/**
	 * @return the properties of the API for the test
	 */
	protected abstract TestProperties getProperties();
}
