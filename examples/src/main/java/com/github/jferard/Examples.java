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

import java.util.logging.Logger;

import com.github.jferard.spreadsheetwrapper.DocumentFactoryManager;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;

public class Examples {
	private Examples() {}
	
	/**
	 * First example.
	 * Needs odfdom in path.
	 * @throws SpreadsheetException
	 */
	public static void createSimpleDocument() throws SpreadsheetException {
		final DocumentFactoryManager manager = new DocumentFactoryManager(Logger.getAnonymousLogger());
		final SpreadsheetDocumentFactory factory = manager.getFactory("ods.odfdom.OdsOdfdomDocumentFactory");
		final SpreadsheetDocumentWriter documentWriter = factory.create();
		final SpreadsheetWriter newSheet = documentWriter.addSheet("0");
		newSheet.setInteger(0, 0, 1);
	}
}
