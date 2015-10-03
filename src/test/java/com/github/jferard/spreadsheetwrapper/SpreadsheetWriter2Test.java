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
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;

import com.github.jferard.spreadsheetwrapper.WrapperCellStyle.Color;

public abstract class SpreadsheetWriter2Test {
	@Rule
	public TestName name = new TestName();
	protected SpreadsheetDocumentFactory factory;
	protected SpreadsheetDocumentWriter sdw;

	protected SpreadsheetWriter sw;

	/** set the test up */
	@Before
	public void setUp() {
		this.factory = this.getFactory();
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
			final File outputFile = SpreadsheetTest.getOutputFile(this
					.getClass().getSimpleName(), this.name.getMethodName(),
					this.getExtension());
			this.sdw.saveAs(outputFile);
			this.sdw.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test
	public void testBoolean() {
		// this.sw.setCellContents(0, 0, true);
		this.sw.setBoolean(0, 0, true);
		Assert.assertEquals(true, this.sw.getBoolean(0, 0));

		this.sw.setBoolean(0, 0, true);
		Assert.assertEquals(true, this.sw.getCellContent(0, 0));
	}

	@Test
	public void testDateDay() {
		final Date d = new Date(52 * 86400000L);
		final Date d2 = this.sw.setDate(2, 2, d);
		Assert.assertEquals(d2, this.sw.getDate(2, 2));

		final Object o2 = this.sw.setCellContent(2, 2, d);
		Assert.assertEquals(d2, this.sw.getCellContent(2, 2));
		Assert.assertEquals(o2, this.sw.getCellContent(2, 2));
	}

	@Test
	public void testDateSecond() {
		final Date d = new Date(5234597000L);
		final Date d2 = this.sw.setDate(2, 2, d);
		Assert.assertEquals(d2, this.sw.getDate(2, 2));

		final Object o2 = this.sw.setCellContent(2, 2, d);
		Assert.assertEquals(d2, this.sw.getCellContent(2, 2));
		Assert.assertEquals(o2, this.sw.getCellContent(2, 2));
	}

	@Test
	public void testDouble() {
		this.sw.setDouble(2, 2, 100.7);
		Assert.assertEquals(Double.valueOf(100.7), this.sw.getDouble(2, 2));

		this.sw.setDouble(2, 2, 100.7);
		Assert.assertEquals(Double.valueOf(100.7), this.sw.getCellContent(2, 2));
	}

	@Test
	public void testFormula() {
		this.sw.setFormula(1, 1, "A1");
		Assert.assertEquals("A1", this.sw.getFormula(1, 1));

		this.sw.setFormula(1, 1, "A1");
		Assert.assertEquals("A1", this.sw.getCellContent(1, 1));
	}

	@Test
	public void testFormula2() {
		this.sw.setFormula(1, 1, "afdferg'|[{*dfgdrsg]");
		Assert.assertEquals("afdferg'|[{*dfgdrsg]", this.sw.getFormula(1, 1));

		this.sw.setFormula(1, 1, "afdferg'|[{*dfgdrsg]");
		Assert.assertEquals("afdferg'|[{*dfgdrsg]",
				this.sw.getCellContent(1, 1));
	}

	@Test
	public void testInteger() {
		this.sw.setInteger(1, 1, 10);
		Assert.assertEquals(Integer.valueOf(10), this.sw.getInteger(1, 1));

		this.sw.setInteger(1, 1, 10);
		Assert.assertEquals(Integer.valueOf(10), this.sw.getCellContent(1, 1));
	}

	@Test
	public final void testStyle() {
		this.sdw.createStyle("mystyle",
				"background-color:#999999;font-weight:bold");
		this.sw.setCellContent(0, 0, "Head", "mystyle");
	}

	@Test
	public final void testStyle2() {
		final WrapperFont wrapperFont = new WrapperFont(true, false, 0, Color.BLACK);
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle(Color.AQUA, wrapperFont);
		this.sdw.setStyle("mystyle2", wrapperCellStyle);
		this.sw.setCellContent(0, 0, "Head", "mystyle2");
		this.sw.setCellContent(1, 0, "Tail", "");
	}

	@Test
	public final void testStyle3() {
		final WrapperFont wrapperFont = new WrapperFont(true, false, 0, Color.BLACK);
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle(Color.AQUA, wrapperFont);
		this.sdw.setStyle("mystyle2", wrapperCellStyle);
		this.sw.setCellContent(0, 0, "Head", "mystyle2");
		this.sw.setCellContent(2, 2, "Tail");
	}

	@Test
	public final void testStyle4() { // buggy for simpleodf
		final WrapperFont wrapperFont = new WrapperFont(true, false, 0, Color.BLACK);
		final WrapperCellStyle wrapperCellStyle = new WrapperCellStyle(Color.AQUA, wrapperFont);
		this.sdw.setStyle("mystyle2", wrapperCellStyle);
		this.sw.setCellContent(0, 0, "Head", "mystyle2");
		this.sw.setCellContent(1, 0, "Tail");
	}

	@Test
	public void testText() {
		this.sw.setText(1, 1, "10");
		Assert.assertEquals("10", this.sw.getText(1, 1));

		this.sw.setText(1, 1, "10");
		Assert.assertEquals("10", this.sw.getCellContent(1, 1));
	}

	@Test
	public void testText1000col() {
		for (int i = 0; i < 1000; i += 100) {
			this.sw.setText(1, i, "10");
			Assert.assertEquals("10", this.sw.getText(1, i));
		}
	}

	@Test
	public void testText1000row() {
		for (int i = 0; i < 1000; i += 100) {
			this.sw.setText(i, 1, "10");
			Assert.assertEquals("10", this.sw.getText(i, 1));
		}
	}

	@Test
	public void testText200col() {
		for (int i = 0; i < 200; i += 50) {
			this.sw.setText(1, i, "100");
			Assert.assertEquals("100", this.sw.getText(1, i));
		}
	}

	@Test(expected = IllegalArgumentException.class)
	public void testTextMinus10a() {
		this.sw.setText(-10, 1, "10");
		Assert.assertEquals("10", this.sw.getText(-10, 1));
	}

	@Test(expected = IllegalArgumentException.class)
	public void testTextMinus10b() {
		this.sw.setText(1, -10, "10");
		Assert.assertEquals("10", this.sw.getText(1, -10));
	}

	protected abstract String getExtension();

	protected abstract SpreadsheetDocumentFactory getFactory();

}
