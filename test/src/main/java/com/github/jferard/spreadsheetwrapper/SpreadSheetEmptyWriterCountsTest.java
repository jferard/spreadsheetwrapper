package com.github.jferard.spreadsheetwrapper;

import java.io.File;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;

public abstract class SpreadSheetEmptyWriterCountsTest {
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
			final File outputFile = SpreadsheetTest.getOutputFile(this
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
	public void testGetIsNotCreateGetThenNothing() {
		this.sw.getCellContent(15, 5);
		Assert.assertEquals(0, this.sw.getRowCount());
	}

	@Test
	public void testGetIsNotCreateGetThenSet() {
		this.sw.getCellContent(15, 5);
		this.sw.setText(1, 3, "a");
		Assert.assertEquals(2, this.sw.getRowCount());
		Assert.assertEquals(4, this.sw.getCellCount(1));
	}

	@Test
	public void testGetIsNotCreateSetThenGet() {
		this.sw.setText(1, 3, "a");
		this.sw.getCellContent(15, 5);
		Assert.assertEquals(2, this.sw.getRowCount());
		Assert.assertEquals(4, this.sw.getCellCount(1));
	}

	@Test(expected = IllegalArgumentException.class)
	public void testNonExistingRowCellCountK11() {
		this.sw.setText(10, 10, "10:10");
		Assert.assertEquals(0, this.sw.getCellCount(11));
	}

	@Test
	public void testRowAndCellCountsK11() {
		this.sw.setText(10, 10, "10:10");
		Assert.assertEquals(11, this.sw.getRowCount());
		Assert.assertEquals(0, this.sw.getCellCount(1));
		Assert.assertEquals(0, this.sw.getCellCount(9));
		Assert.assertEquals(11, this.sw.getCellCount(10));
	}

	@Test
	public void testRowCountA1() {
		this.sw.setInteger(0, 0, 1);
		Assert.assertEquals(1, this.sw.getRowCount());
		Assert.assertEquals(1, this.sw.getCellCount(0));
	}

	@Test
	public void testRowCountD2() {
		this.sw.setInteger(1, 3, 1);
		Assert.assertEquals(2, this.sw.getRowCount());
		Assert.assertEquals(4, this.sw.getCellCount(1));
	}

	@Test
	public void testRowCountZero() {
		Assert.assertEquals(0, this.sw.getRowCount());
	}

	protected abstract TestProperties getProperties();

}
