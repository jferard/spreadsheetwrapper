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

import java.util.logging.Logger;

import org.junit.Test;

import com.github.jferard.spreadsheetwrapper.ods.odfdom.OdsOdfdomDocumentFactory;
import com.github.jferard.spreadsheetwrapper.ods.simpleodf.OdsSimpleodfDocumentFactory;
import com.github.jferard.spreadsheetwrapper.ods.simpleods.OdsSimpleodsDocumentFactory;
import com.github.jferard.spreadsheetwrapper.xls.jxl.XlsJxlDocumentFactory;
import com.github.jferard.spreadsheetwrapper.xls.poi.XlsPoiDocumentFactory;

@SuppressWarnings("static-method")
public class DocumentFactoryManagerTest {
	@Test
	public final void testDirect() {
		final Logger logger = null;
		@SuppressWarnings("unused")
		SpreadsheetDocumentFactory factory;
		factory = OdsSimpleodsDocumentFactory.create(logger);
		factory = OdsSimpleodfDocumentFactory.create(logger);
		factory = OdsOdfdomDocumentFactory.create(logger);
		factory = com.github.jferard.spreadsheetwrapper.ods.jopendocument12.OdsJOpenDocumentFactory
				.create(logger);

		factory = XlsJxlDocumentFactory.create(logger);
		factory = XlsPoiDocumentFactory.create(logger);
	}

	@Test
	public final void testManager() throws SpreadsheetException,
			IllegalArgumentException {
		final DocumentFactoryManager manager = new DocumentFactoryManager(null);
		@SuppressWarnings("unused")
		SpreadsheetDocumentFactory factory;
		factory = manager
				.getFactory("ods.simpleodf.OdsSimpleodfDocumentFactory");
		factory = manager
				.getFactory("ods.simpleods.OdsSimpleodsDocumentFactory");
		factory = manager.getFactory("ods.odfdom.OdsOdfdomDocumentFactory");
		factory = manager
				.getFactory("ods.jopendocument12.OdsJOpenDocumentFactory");

		factory = manager.getFactory("xls.jxl.XlsJxlDocumentFactory");
		factory = manager.getFactory("xls.poi.XlsPoiDocumentFactory");

	}

}
