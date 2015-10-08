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
package com.github.jferard.spreadsheetwrapper.impl;

import java.io.IOException;
import java.io.InputStream;
import java.util.logging.Logger;

import org.junit.Assert;
import org.junit.Before;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyleHelper;
import com.github.jferard.spreadsheetwrapper.ods.odfdom.OdsOdfdomDocumentFactory;
import com.github.jferard.spreadsheetwrapper.ods.odfdom.OdsOdfdomStyleUtility;

public class SpreadsheetWriterCursorImplTest extends CursorAbstractTest {
	@Before
	public void setUp() {
		final SpreadsheetDocumentFactory factory = this.getFactory();
		try {
			final InputStream inputStream = this
					.getClass()
					.getResource(
							String.format("/VilleMTP_MTP_MonumentsHist.%s",
									this.getExtension())).openStream();
			final SpreadsheetDocumentWriter sdw = factory
					.openForWrite(inputStream);
			final SpreadsheetWriter sheet = sdw.getSpreadsheet(0);
			this.rowCount = sheet.getRowCount();
			this.colCount = sheet.getCellCount(0);
			this.cursor = sheet.getNewCursor();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail();
		} catch (final IOException e) {
			e.printStackTrace();
			Assert.fail();
		}
	}

	protected String getExtension() {
		return "ods";
	}

	protected SpreadsheetDocumentFactory getFactory() {
		return new OdsOdfdomDocumentFactory(Logger.getGlobal(),
				new OdsOdfdomStyleUtility(new WrapperCellStyleHelper()));
	}

}
