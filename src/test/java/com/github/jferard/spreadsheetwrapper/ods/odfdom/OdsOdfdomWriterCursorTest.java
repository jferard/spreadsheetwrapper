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
package com.github.jferard.spreadsheetwrapper.ods.odfdom;

import java.util.logging.Logger;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursorTest;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyleHelper;

public class OdsOdfdomWriterCursorTest extends SpreadsheetWriterCursorTest {
	// @Rule
	// public TestName name = new TestName();
	//
	// /** {@inheritDoc} */
	// @Override
	// @Before
	// public void setUp() {
	// this.factory = getFactory();
	// try {
	// final InputStream inputStream = this
	// .getClass()
	// .getResource(
	// String.format("/VilleMTP_MTP_MonumentsHist.%s",
	// this.getExtension())).openStream();
	// this.sdw = this.factory.openCopyOf(inputStream);
	// Assert.assertEquals(1, this.sdw.getSheetCount());
	// this.sw = this.sdw.getSpreadsheet(0);
	// this.swc = this.sw.getNewCursor();
	// } catch (final SpreadsheetException e) {
	// e.printStackTrace();
	// Assert.fail();
	// } catch (IOException e) {
	// e.printStackTrace();
	// Assert.fail();
	// }
	// }

	@Override
	protected String getExtension() {
		return "ods";
	}

	@Override
	protected SpreadsheetDocumentFactory getFactory() {
		return new OdsOdfdomDocumentFactory(Logger.getGlobal(), new OdsOdfdomStyleUtility(new WrapperCellStyleHelper()));
	}
}
