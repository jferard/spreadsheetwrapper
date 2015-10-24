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
package com.github.jferard;

import org.junit.Assert;
import org.junit.Test;

import com.github.jferard.spreadsheetwrapper.DocumentFactoryManager;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;

/**
 * Unit test for simple App.
 */
public class AppTest {
	/**
	 * Rigourous Test :-)
	 *
	 * @throws SpreadsheetException
	 */
	@Test
	public void testApp() throws SpreadsheetException {
		final DocumentFactoryManager manager = new DocumentFactoryManager(null);
		SpreadsheetDocumentFactory factory;
		factory = manager.getFactory("ods.odfdom.OdsOdfdomDocumentFactory");
		final SpreadsheetDocumentWriter documentWriter = factory.create();
		final SpreadsheetWriter newSheet = documentWriter.addSheet("0");
		newSheet.setInteger(0, 0, 1);
		Assert.assertEquals(1, newSheet.getRowCount());
		Assert.assertEquals(1, newSheet.getCellCount(0));
	}
}
