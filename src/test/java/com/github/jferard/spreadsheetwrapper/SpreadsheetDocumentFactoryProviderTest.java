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
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Logger;

import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import com.github.jferard.spreadsheetwrapper.ods.odfdom.OdsOdfdomDocumentFactory;

@SuppressWarnings("unused")
public class SpreadsheetDocumentFactoryProviderTest {

	private SpreadsheetDocumentFactoryProvider spreadsheetDocumentFactoryProvider;

	@Before
	public void setUp() {
		final Map<String, SpreadsheetDocumentFactory> factoryByExtension = new HashMap<String, SpreadsheetDocumentFactory>();
		factoryByExtension.put("ods",
				new OdsOdfdomDocumentFactory(Logger.getGlobal()));
		this.spreadsheetDocumentFactoryProvider = new SpreadsheetDocumentFactoryProvider(
				factoryByExtension);
	}

	@Test(expected = IllegalArgumentException.class)
	public final void test() {
		final SpreadsheetDocumentFactory factory = this.spreadsheetDocumentFactoryProvider
				.getFactory("xls");
	}

	@Test(expected = IllegalStateException.class)
	public final void test2() {
		final SpreadsheetDocumentFactory factory = this.spreadsheetDocumentFactoryProvider
				.getFactory("ods");
		try {
			final SpreadsheetDocumentWriter s = factory.create();
			s.save();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	@Test
	public final void testCreateAndSaveAs() {
		this.testCreateAndSaveAs("ods");
		this.testCreateAndSaveAs("xls");
	}

	private final void testCreateAndSaveAs(final String extension) {
		final SpreadsheetDocumentFactory factory = this.spreadsheetDocumentFactoryProvider
				.getFactory("ods");
		try {
			final SpreadsheetDocumentWriter s = factory.create();
			s.saveAs(new File(System.getProperty("java.io.tmpdir"), String
					.format("test.%s", extension)));
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}
}
