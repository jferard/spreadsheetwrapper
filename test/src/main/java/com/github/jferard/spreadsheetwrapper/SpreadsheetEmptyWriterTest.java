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

import org.junit.After;
import org.junit.Assert;
import org.junit.Assume;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;

import com.github.jferard.spreadsheetwrapper.impl.StyleUtility;

public abstract class SpreadsheetEmptyWriterTest {
	@Rule
	public TestName name = new TestName();
	protected SpreadsheetDocumentFactory factory;
	protected SpreadsheetDocumentWriter sdw;

	protected SpreadsheetWriter sw;

	/** set the test up */
	@Before
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			this.sdw = this.factory.create();
			this.sw = this.sdw.addSheet(0, "first sheet");
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
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
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test
	public final void testCreate2StylesFromObjectsAndSetStylesWithGap() {
		final WrapperFont wrapperFont = new WrapperFont().setBold();
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle(
				WrapperColor.AQUA, wrapperFont);
		this.sdw.setStyle("mystyle2", wrapperCellStyle);
		this.sw.setCellContent(0, 0, "Head", "mystyle2");
		this.sw.setCellContent(2, 2, "Tail");

		try {
			final WrapperCellStyle newCellStyle = this.sw.getStyle(0, 0);
			Assert.assertEquals(wrapperCellStyle, newCellStyle);
			WrapperCellStyle otherCellStyle = this.sw.getStyle(1, 0);
			Assert.assertNotEquals(wrapperCellStyle, otherCellStyle);
			otherCellStyle = this.sw.getStyle(2, 0);
			// Assert.assertNotEquals(wrapperCellStyle, otherCellStyle);
			otherCellStyle = this.sw.getStyle(0, 1);
			Assert.assertNotEquals(wrapperCellStyle, otherCellStyle);
			otherCellStyle = this.sw.getStyle(0, 2);
			// Assert.assertNotEquals(wrapperCellStyle, otherCellStyle);
			otherCellStyle = this.sw.getStyle(1, 1);
			// Assert.assertNotEquals(wrapperCellStyle, otherCellStyle);
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public final void testCreate2StylesFromObjectsAndSetStylesWithNoGap() { // buggy
		// for
		// simpleodf
		final WrapperFont wrapperFont = new WrapperFont().setBold();
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle(
				WrapperColor.AQUA, wrapperFont);
		this.sdw.setStyle("mystyle2", wrapperCellStyle);
		this.sw.setCellContent(0, 0, "Head", "mystyle2");
		this.sw.setCellContent(1, 0, "Tail");

		try {
			final WrapperCellStyle newCellStyle = this.sw.getStyle(0, 0);
			Assert.assertEquals(wrapperCellStyle, newCellStyle);
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public final void testCreateStyleFromObjectsAndSetStyle() {
		final WrapperFont wrapperFont = new WrapperFont().setBold();
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle(
				WrapperColor.AQUA, wrapperFont);
		this.sdw.setStyle("mystyle2", wrapperCellStyle);
		this.sw.setCellContent(0, 0, "Head", "mystyle2");
		this.sw.setCellContent(1, 0, "Tail", "");

		try {
			final WrapperCellStyle newCellStyle = this.sw.getStyle(0, 0);
			Assert.assertEquals(wrapperCellStyle, newCellStyle);
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public final void testCreateStyleFromStringAndSetStyle() {
		final StyleUtility utility = new StyleUtility();
		final WrapperCellStyle wrapperCellStyle = utility
				.getWrapperCellStyle("background-color:#999999;font-weight:bold");
		this.sdw.setStyle("mystyle", wrapperCellStyle);
		this.sw.setCellContent(0, 0, "Head", "mystyle");

		try {
			final WrapperCellStyle newCellStyle = this.sw.getStyle(0, 0);
			Assert.assertEquals(wrapperCellStyle, newCellStyle);
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public void testSetBoolean() {
		try {
			this.sw.setCellContent(0, 0, true);
			Assert.assertEquals(true, this.sw.getBoolean(0, 0));

			this.sw.setBoolean(0, 0, true);
			Assert.assertEquals(true, this.sw.getCellContent(0, 0));
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public void testSetDate() {
		final Date dd = new Date(52 * 86400000L);
		final Date ds = new Date(52 * 86400000L + 12345000L);
		final Date dm = new Date(52 * 86400000L + 12345789L);

		final Date d2 = this.sw.setDate(2, 2, dm);
		Assert.assertTrue(d2.equals(dm) || d2.equals(ds) || d2.equals(dd));
		Assert.assertEquals(d2, this.sw.getDate(2, 2));
		Assert.assertEquals(d2, this.sw.getCellContent(2, 2));

		final Object o2 = this.sw.setCellContent(2, 2, dm);
		Assert.assertTrue(o2.equals(dm) || o2.equals(ds) || o2.equals(dd));
		Assert.assertEquals(o2, this.sw.getDate(2, 2));
		Assert.assertEquals(o2, this.sw.getCellContent(2, 2));
	}

	@Test
	public void testSetDouble() {
		this.sw.setCellContent(2, 2, 100.7);
		Assert.assertEquals(Double.valueOf(100.7), this.sw.getDouble(2, 2));

		this.sw.setDouble(2, 2, 100.7);
		Assert.assertEquals(Double.valueOf(100.7), this.sw.getCellContent(2, 2));
	}

	@Test
	public void testSetFormula() {
		try {
			this.sw.setFormula(1, 1, "A1");
			Assert.assertEquals("A1", this.sw.getFormula(1, 1));

			this.sw.setFormula(1, 1, "A1");
			Assert.assertEquals("A1", this.sw.getCellContent(1, 1));
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public void testSetIncorrrectFormula() {
		try {
			this.sw.setFormula(1, 1, "afdferg'|[{*dfgdrsg]");
			Assert.assertEquals("afdferg'|[{*dfgdrsg]",
					this.sw.getFormula(1, 1));

			this.sw.setFormula(1, 1, "afdferg'|[{*dfgdrsg]");
			Assert.assertEquals("afdferg'|[{*dfgdrsg]",
					this.sw.getCellContent(1, 1));
		} catch (final IllegalArgumentException e) {
			Assume.assumeNoException(e);
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public void testSetInteger() {
		this.sw.setCellContent(1, 1, 10);
		Assert.assertEquals(Integer.valueOf(10), this.sw.getInteger(1, 1));

		this.sw.setInteger(1, 1, 10);
		Assert.assertEquals(Integer.valueOf(10), this.sw.getCellContent(1, 1));
	}

	@Test
	public void testSetText() {
		this.sw.setCellContent(1, 1, "10");
		Assert.assertEquals("10", this.sw.getText(1, 1));

		this.sw.setText(1, 1, "10");
		Assert.assertEquals("10", this.sw.getCellContent(1, 1));
	}

	@Test
	public void testSetTextOnCol0To1000() {
		int i = 0;
		try {
			for (i = 0; i < 1000; i += 100) {
				this.sw.setText(1, i, "10");
				Assert.assertEquals("10", this.sw.getText(1, i));
			}
		} catch (final UnsupportedOperationException e) {
			if (i >= 255)
				Assume.assumeNoException(e);
			else
				throw e;
		}
	}

	@Test
	public void testSetTextOnCol255() {
		this.sw.setText(1, 255, "100");
		Assert.assertEquals("100", this.sw.getText(1, 255));
	}

	@Test(expected = IllegalArgumentException.class)
	public void testSetTextOnColMinus10() {
		this.sw.setText(1, -10, "10");
		Assert.assertEquals("10", this.sw.getText(1, -10));
	}

	@Test
	public void testSetTextOnRow0To100000() {
		int j = 0;
		try {
			for (j = 0; j < 100000; j += 20000) {
				this.sw.setText(j, 1, "10");
				Assert.assertEquals("10", this.sw.getText(j, 1));
			}
		} catch (final UnsupportedOperationException e) {
			if (j >= 65535)
				Assume.assumeNoException(e);
			else
				throw e;
		}
	}

	@Test
	public void testSetTextOnRow50() {
		this.sw.setText(50, 1, "10");
		Assert.assertEquals("10", this.sw.getText(50, 1));
	}

	@Test
	public void testSetTextOnRow65535() {
		this.sw.setText(0, 1, "0");
		this.sw.setText(65535, 1, "100");
		Assert.assertEquals("100", this.sw.getText(65535, 1));
	}

	@Test(expected = IllegalArgumentException.class)
	public void testSetTextOnRowMinus10() {
		this.sw.setText(-10, 1, "10");
		Assert.assertEquals("10", this.sw.getText(-10, 1));
	}

	protected abstract TestProperties getProperties();
}
