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
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.ListIterator;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.junit.After;
import org.junit.Assert;
import org.junit.Assume;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;

import com.github.jferard.spreadsheetwrapper.Cursor.Move;
import com.google.common.collect.Lists;

public abstract class AbstractSpreadsheetWriterCursorTest {
	/** last data row for tests */
	private static final List<Object> LAST_ROW = Lists
			.<Object> newArrayList(
					Integer.valueOf(1201),
					"Monuments historiques classés ou inscrits",
					"Collège des Ecossais",
					"Inscrit Inv. MH 19 12 2013",
					"Ensemble de la parcelle avec terrasses et jardins, façades et toitures des bâtiments, la tour dite \"outlook tower\" en totalité et le monument à Jeanne d'Arc du collège",
					Double.valueOf(35285.3539), Double.valueOf(943.342686614));

	/** head for tests */
	private static final List<Object> ROW_1 = Lists.<Object> newArrayList(
			"OBJECTID,N,10,0", "LIBELLE,C,254", "ETIQUETTE,C,254",
			"DATE_CREAT,C,254", "COMMENTAIR,C,254", "SHAPE_AREA,N,19,11",
			"SHAPE_LEN,N,19,11");

	/** first data row for tests */
	private static final List<Object> ROW_2 = Lists.<Object> newArrayList(
			Integer.valueOf(170), "Monuments historiques classés ou inscrits",
			"Immeuble dit « Hôtel Lefèvre »", "inscrit Inv. MH 19 11 1985",
			"En totalité", Double.valueOf(333.810638315),
			Double.valueOf(93.4774612976));

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

	/** the cursor */
	protected SpreadsheetWriterCursor sheetWriterCursor;

	/** set the test up */
	@Before
	@SuppressWarnings("nullness")
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			final URL sourceURL = this.getProperties().getResourceURL();
			Assume.assumeNotNull(sourceURL);

			final InputStream inputStream = sourceURL.openStream();
			final File outputFile = SpreadsheetTestHelper.getOutputFile(
					this.factory, this.getClass().getSimpleName(),
					this.name.getMethodName());
			final OutputStream outputStream = new FileOutputStream(outputFile);
			this.documentWriter = this.factory.openForWrite(inputStream,
					outputStream);
			Assert.assertEquals(1, this.documentWriter.getSheetCount());
			this.sheetWriter = this.documentWriter.getSpreadsheet(0);
			this.sheetWriterCursor = this.sheetWriter.getNewCursor();
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
	public void tearDown() {
		try {
			this.documentWriter.save();
			this.documentWriter.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** test a circular move */
	@Test
	public void testCircularMove() {
		final int r = 5;
		final int c = 4;
		Assert.assertEquals(Move.JUMP, this.sheetWriterCursor.set(r, c));
		Assert.assertEquals(r, this.sheetWriterCursor.getR());
		Assert.assertEquals(c, this.sheetWriterCursor.getC());
		this.sheetWriterCursor.right();
		Assert.assertEquals(r, this.sheetWriterCursor.getR());
		Assert.assertEquals(c + 1, this.sheetWriterCursor.getC());
		this.sheetWriterCursor.up();
		Assert.assertEquals(r - 1, this.sheetWriterCursor.getR());
		Assert.assertEquals(c + 1, this.sheetWriterCursor.getC());
		this.sheetWriterCursor.left();
		Assert.assertEquals(r - 1, this.sheetWriterCursor.getR());
		Assert.assertEquals(c, this.sheetWriterCursor.getC());
		this.sheetWriterCursor.left();
		Assert.assertEquals(r - 1, this.sheetWriterCursor.getR());
		Assert.assertEquals(c - 1, this.sheetWriterCursor.getC());
		this.sheetWriterCursor.down();
		Assert.assertEquals(r, this.sheetWriterCursor.getR());
		Assert.assertEquals(c - 1, this.sheetWriterCursor.getC());
		this.sheetWriterCursor.right();
		Assert.assertEquals(r, this.sheetWriterCursor.getR());
		Assert.assertEquals(c, this.sheetWriterCursor.getC());
	}

	/** test simple forward move from first cell */
	@Test
	public final void testMoveFromFirstRow() {
		this.sheetWriterCursor.first();
		final ArrayList<Object> list = Lists.<Object> newArrayList();
		list.addAll(AbstractSpreadsheetWriterCursorTest.ROW_1);
		list.addAll(AbstractSpreadsheetWriterCursorTest.ROW_2);
		final ListIterator<Object> iterator = list.listIterator();

		final int COL_COUNT = AbstractSpreadsheetWriterCursorTest.ROW_1.size();

		try {
			while (iterator.hasNext()) {
				final int index = iterator.nextIndex();
				Assert.assertEquals(index / COL_COUNT,
						this.sheetWriterCursor.getR());
				Assert.assertEquals(index % COL_COUNT,
						this.sheetWriterCursor.getC());
				Assert.assertEquals(iterator.next(),
						this.sheetWriterCursor.getCellContent());
				if (index % COL_COUNT == COL_COUNT - 1)
					Assert.assertEquals(Move.JUMP,
							this.sheetWriterCursor.next());
				else
					Assert.assertEquals(Move.NEXT,
							this.sheetWriterCursor.next());
			}
		} catch (final IllegalArgumentException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** test simple forward move from first cell of last row */
	@Test
	public final void testMoveFromLastRow() {
		final int lastRow = this.sheetWriter.getRowCount() - 1;
		this.sheetWriterCursor.set(lastRow, 0);
		final ListIterator<Object> iterator = AbstractSpreadsheetWriterCursorTest.LAST_ROW
				.listIterator();

		final int COL_COUNT = AbstractSpreadsheetWriterCursorTest.ROW_1.size();

		try {
			while (iterator.hasNext()) {
				final int index = iterator.nextIndex();
				Assert.assertEquals(lastRow + index / COL_COUNT,
						this.sheetWriterCursor.getR());
				Assert.assertEquals(index % COL_COUNT,
						this.sheetWriterCursor.getC());
				Assert.assertEquals(iterator.next(),
						this.sheetWriterCursor.getCellContent());
				if (index % COL_COUNT == COL_COUNT - 1)
					Assert.assertEquals(Move.FAIL,
							this.sheetWriterCursor.next());
				else
					Assert.assertEquals(Move.NEXT,
							this.sheetWriterCursor.next());
			}
			Assert.assertEquals(
					AbstractSpreadsheetWriterCursorTest.LAST_ROW
							.get(AbstractSpreadsheetWriterCursorTest.LAST_ROW
									.size() - 1), this.sheetWriterCursor
							.getCellContent());
		} catch (final IllegalArgumentException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail();
		}
	}

	/** test set a value */
	@Test
	public void testSet() {
		final int r = 5;
		final int c = 4;
		Assert.assertEquals(Move.JUMP, this.sheetWriterCursor.set(r, c));
		try {
			this.sheetWriterCursor.setDate(new Date(0)); // setDate : 0 UTC = 1
			// CET
			Assert.assertEquals(new Date(0),
					this.sheetWriterCursor.getCellContent());
			Assert.assertEquals(new Date(0), this.sheetWriterCursor.getDate()); // getDate
			// :
			// 0
			// CET = 0
			// CET
		} catch (final IllegalArgumentException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		}
	}

	/** test set out of range. Should set to 0, 0 */
	@Test
	public void testSetOutOfRange() {
		Assert.assertEquals(Move.FAIL, this.sheetWriterCursor.set(10, 10));
		Assert.assertEquals(0, this.sheetWriterCursor.getR());
		Assert.assertEquals(0, this.sheetWriterCursor.getC());
	}

	/** the properties for test */
	protected abstract TestProperties getProperties();
}
