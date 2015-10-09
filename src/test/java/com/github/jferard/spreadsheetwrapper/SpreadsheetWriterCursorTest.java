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
import java.util.ArrayList;
import java.util.Date;
import java.util.ListIterator;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;

import com.github.jferard.spreadsheetwrapper.Cursor.Move;
import com.google.common.collect.Lists;

public abstract class SpreadsheetWriterCursorTest {
	@Rule
	public TestName name = new TestName();
	protected SpreadsheetDocumentFactory factory;
	protected SpreadsheetDocumentWriter sdw;
	protected SpreadsheetWriter sw;

	protected SpreadsheetWriterCursor swc;

	/** set the test up */
	@Before
	public void setUp() {
		this.factory = this.getProperties().getFactory();
		try {
			final InputStream inputStream = this
					.getClass()
					.getResource(
							String.format("/VilleMTP_MTP_MonumentsHist.%s",
									this.getProperties().getExtension())).openStream();
			final File outputFile = SpreadsheetTest.getOutputFile(this
					.getClass().getSimpleName(), this.name.getMethodName(),
					this.getProperties().getExtension());
			final OutputStream outputStream = new FileOutputStream(outputFile);
			this.sdw = this.factory.openForWrite(inputStream, outputStream);
			Assert.assertEquals(1, this.sdw.getSheetCount());
			this.sw = this.sdw.getSpreadsheet(0);
			this.swc = this.sw.getNewCursor();
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
	public void tearDown() {
		try {
			this.sdw.save();
			this.sdw.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test
	public final void testMove1() {
		this.swc.first();
		final ArrayList<Object> list = Lists.<Object> newArrayList(
				"OBJECTID,N,10,0", "LIBELLE,C,254", "ETIQUETTE,C,254",
				"DATE_CREAT,C,254", "COMMENTAIR,C,254", "SHAPE_AREA,N,19,11",
				"SHAPE_LEN,N,19,11", Integer.valueOf(170),
				"Monuments historiques classés ou inscrits",
				"Immeuble dit « Hôtel Lefèvre »", "inscrit Inv. MH 19 11 1985",
				"En totalité", Double.valueOf(333.810638315),
				Double.valueOf(93.4774612976));
		final ListIterator<Object> it = list.listIterator();

		try {
			while (it.hasNext()) {
				final int n = it.nextIndex();
				Assert.assertEquals(n / 7, this.swc.getR());
				Assert.assertEquals(n % 7, this.swc.getC());
				Assert.assertEquals(it.next(), this.swc.getCellContent());
				if (n % 7 == 6)
					Assert.assertEquals(Move.JUMP, this.swc.next());
				else
					Assert.assertEquals(Move.NEXT, this.swc.next());
			}
		} catch (final IllegalArgumentException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test
	public final void testMove2() {
		final int lastRow = this.sw.getRowCount() - 1;
		this.swc.set(lastRow, 0);
		final ArrayList<Object> list = Lists
				.<Object> newArrayList(
						Integer.valueOf(1201),
						"Monuments historiques classés ou inscrits",
						"Collège des Ecossais",
						"Inscrit Inv. MH 19 12 2013",
						"Ensemble de la parcelle avec terrasses et jardins, façades et toitures des bâtiments, la tour dite \"outlook tower\" en totalité et le monument à Jeanne d'Arc du collège",
						Double.valueOf(35285.3539),
						Double.valueOf(943.342686614));
		final ListIterator<Object> it = list.listIterator();

		try {
			while (it.hasNext()) {
				final int n = it.nextIndex();
				Assert.assertEquals(lastRow + n / 7, this.swc.getR());
				Assert.assertEquals(n % 7, this.swc.getC());
				Assert.assertEquals(it.next(), this.swc.getCellContent());
				if (n % 7 == 6)
					Assert.assertEquals(Move.FAIL, this.swc.next());
				else
					Assert.assertEquals(Move.NEXT, this.swc.next());
			}
			Assert.assertEquals(list.get(list.size() - 1),
					this.swc.getCellContent());
		} catch (final IllegalArgumentException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test
	public void testMove3() {
		Assert.assertEquals(Move.FAIL, this.swc.set(10, 10));
		Assert.assertEquals(0, this.swc.getR());
		Assert.assertEquals(0, this.swc.getC());
	}

	@Test
	public void testMove4() {
		final int r = 5;
		final int c = 4;
		Assert.assertEquals(Move.JUMP, this.swc.set(r, c));
		Assert.assertEquals(r, this.swc.getR());
		Assert.assertEquals(c, this.swc.getC());
		this.swc.right();
		Assert.assertEquals(r, this.swc.getR());
		Assert.assertEquals(c + 1, this.swc.getC());
		this.swc.up();
		Assert.assertEquals(r - 1, this.swc.getR());
		Assert.assertEquals(c + 1, this.swc.getC());
		this.swc.left();
		Assert.assertEquals(r - 1, this.swc.getR());
		Assert.assertEquals(c, this.swc.getC());
		this.swc.left();
		Assert.assertEquals(r - 1, this.swc.getR());
		Assert.assertEquals(c - 1, this.swc.getC());
		this.swc.down();
		Assert.assertEquals(r, this.swc.getR());
		Assert.assertEquals(c - 1, this.swc.getC());
		this.swc.right();
		Assert.assertEquals(r, this.swc.getR());
		Assert.assertEquals(c, this.swc.getC());
	}

	@Test
	public void testSet() {
		final int r = 5;
		final int c = 4;
		Assert.assertEquals(Move.JUMP, this.swc.set(r, c));
		try {
			this.swc.setDate(new Date(0)); // setDate : 0 UTC = 1 CET
			Assert.assertEquals(new Date(0), this.swc.getCellContent());
			Assert.assertEquals(new Date(0), this.swc.getDate()); // getDate : 0
			// CET = 0
			// CET
		} catch (final IllegalArgumentException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		}
	}

	protected abstract TestProperties getProperties();
}
