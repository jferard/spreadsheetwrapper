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
package com.github.jferard.spreadsheetwrapper.ods.jopendocument;

import java.util.logging.Logger;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursorTest;

public class OdsJOpenWriterCursorTest extends SpreadsheetWriterCursorTest {
	@Override
	protected String getExtension() {
		return "ods";
	}

	@Override
	protected SpreadsheetDocumentFactory getFactory() {
		return new OdsJOpenDocumentFactory(Logger.getGlobal());
	}
}