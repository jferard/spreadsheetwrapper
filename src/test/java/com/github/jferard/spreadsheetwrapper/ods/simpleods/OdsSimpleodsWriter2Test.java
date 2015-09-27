/*******************************************************************************
 *     SpreadsheetWrapper - An abstraction layer over the API for Excel or Calc
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
package com.github.jferard.spreadsheetwrapper.ods.simpleods;

import java.io.File;
import java.util.logging.Logger;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetTest;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter2Test;

public class OdsSimpleodsWriter2Test extends SpreadsheetWriter2Test {
	/** set the test up */
	@Before
	@Override
	public void setUp() {
		this.factory = this.getFactory();
		try {
			final File outputFile = SpreadsheetTest.getOutputFile(this
					.getClass().getSimpleName(), this.name.getMethodName(),
					this.getExtension());
			this.sdw = this.factory.create(outputFile);
			this.sw = this.sdw.addSheet(0, "name");
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	/** tear the test down */
	@After
	@Override
	public void tearDown() {
		try {
			this.sdw.save();
			this.sdw.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	/** {@inheritDoc} */
	@Override
	@Test
	public void testBoolean() {
		//
	}

	/** {@inheritDoc} */
	@Override
	@Test
	public void testFormula() {
		//
	}

	/** {@inheritDoc} */
	@Override
	@Test
	public void testFormula2() {
		//
	}

	@Override
	protected String getExtension() {
		return "ods";
	}

	@Override
	protected SpreadsheetDocumentFactory getFactory() {
		return new OdsSimpleodsDocumentFactory(Logger.getGlobal());
	}

	@Override
	public void testDateSecond() {
		// 
	}
}
