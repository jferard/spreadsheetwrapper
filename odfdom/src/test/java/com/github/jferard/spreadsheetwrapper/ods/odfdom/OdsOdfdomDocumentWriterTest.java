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

import org.junit.Test;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriterTest;
import com.github.jferard.spreadsheetwrapper.TestProperties;

public class OdsOdfdomDocumentWriterTest extends SpreadsheetDocumentWriterTest {
	/** {@inheritDoc} */
	@Override
	@Test(expected = UnsupportedOperationException.class)
	public final void testAddAt0() throws IndexOutOfBoundsException,
	CantInsertElementInSpreadsheetException {
		super.testAddAt0();
	}

	/** {@inheritDoc} */
	@Override
	@Test(expected = UnsupportedOperationException.class)
	public final void testAppendAdd20At0()
			throws CantInsertElementInSpreadsheetException {
		super.testAppendAdd20At0();
	}

	@Override
	protected TestProperties getProperties() {
		return OdsOdfdomTestProperties.getProperties();
	}
}
