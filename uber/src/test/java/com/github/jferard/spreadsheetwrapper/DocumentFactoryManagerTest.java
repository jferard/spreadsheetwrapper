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

import org.junit.Before;
import org.junit.Test;
import org.powermock.api.easymock.PowerMock;

import com.github.jferard.spreadsheetwrapper.ods.odfdom.OdsOdfdomDocumentFactory;
import com.github.jferard.spreadsheetwrapper.ods.simpleodf.OdsSimpleodfDocumentFactory;
import com.github.jferard.spreadsheetwrapper.ods.simpleods.OdsSimpleodsDocumentFactory;
import com.github.jferard.spreadsheetwrapper.xls.jxl.XlsJxlDocumentFactory;
import com.github.jferard.spreadsheetwrapper.xls.poi.XlsPoiDocumentFactory;

public class DocumentFactoryManagerTest {
	/** the logger */
	private Logger logger;

	/** create the mock for the logger */
	@Before
	public void setUp() {
		this.logger = PowerMock.createNiceMock(Logger.class);
	}

	/** the the creation of the factories without the manager */
	@Test
	@SuppressWarnings({ "unused", "PMD.DataflowAnomalyAnalysis" })
	public final void testDirect() {
		SpreadsheetDocumentFactory factory;
		factory = OdsSimpleodsDocumentFactory.create(this.logger);
		factory = OdsSimpleodfDocumentFactory.create(this.logger);
		factory = OdsOdfdomDocumentFactory.create(this.logger);
		factory = com.github.jferard.spreadsheetwrapper.ods.jopendocument1_2.OdsJOpenDocumentFactory
				.create(this.logger);
		factory = com.github.jferard.spreadsheetwrapper.ods.jopendocument1_3.OdsJOpenDocumentFactory
				.create(this.logger);

		factory = XlsJxlDocumentFactory.create(this.logger);
		factory = XlsPoiDocumentFactory.create(this.logger);
	}

	/** the the creation of the factories with thee manager */
	@Test
	@SuppressWarnings({ "unused", "PMD.DataflowAnomalyAnalysis" })
	public final void testManager() throws SpreadsheetException,
	IllegalArgumentException {
		final DocumentFactoryManager manager = new DocumentFactoryManager(
				this.logger);
		SpreadsheetDocumentFactory factory;
		factory = manager
				.getFactory("ods.simpleodf.OdsSimpleodfDocumentFactory");
		factory = manager
				.getFactory("ods.simpleods.OdsSimpleodsDocumentFactory");
		factory = manager.getFactory("ods.odfdom.OdsOdfdomDocumentFactory");
		factory = manager
				.getFactory("ods.jopendocument1_2.OdsJOpenDocumentFactory");
		factory = manager
				.getFactory("ods.jopendocument1_3.OdsJOpenDocumentFactory");

		factory = manager.getFactory("xls.jxl.XlsJxlDocumentFactory");
		factory = manager.getFactory("xls.poi.XlsPoiDocumentFactory");
	}

}
