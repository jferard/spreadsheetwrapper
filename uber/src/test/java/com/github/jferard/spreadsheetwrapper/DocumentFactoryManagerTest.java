package com.github.jferard.spreadsheetwrapper;

import java.lang.reflect.InvocationTargetException;
import java.util.logging.Logger;

import org.junit.Test;

import com.github.jferard.spreadsheetwrapper.ods.odfdom.OdsOdfdomDocumentFactory;
import com.github.jferard.spreadsheetwrapper.ods.simpleodf.OdsSimpleodfDocumentFactory;
import com.github.jferard.spreadsheetwrapper.ods.simpleods.OdsSimpleodsDocumentFactory;
import com.github.jferard.spreadsheetwrapper.xls.jxl.XlsJxlDocumentFactory;
import com.github.jferard.spreadsheetwrapper.xls.poi.XlsPoiDocumentFactory;

public class DocumentFactoryManagerTest {
	@Test
	public final void testDirect()  {
		Logger logger = null;
		SpreadsheetDocumentFactory factory;
		factory = OdsSimpleodsDocumentFactory.create(logger);
		factory = OdsSimpleodfDocumentFactory.create(logger);
		factory = OdsOdfdomDocumentFactory.create(logger);
		factory = com.github.jferard.spreadsheetwrapper.ods.jopendocument12.OdsJOpenDocumentFactory.create(logger);
		
		factory = XlsJxlDocumentFactory.create(logger);
		factory = XlsPoiDocumentFactory.create(logger);
	}
	
	@Test
	public final void testManager() throws ClassNotFoundException, InstantiationException, IllegalAccessException, SpreadsheetException, IllegalArgumentException, InvocationTargetException {
		DocumentFactoryManager manager = new DocumentFactoryManager(null);
		SpreadsheetDocumentFactory factory;
		factory = manager.getFactory("ods.simpleodf.OdsSimpleodfDocumentFactory");
		factory = manager.getFactory("ods.simpleods.OdsSimpleodsDocumentFactory");
		factory = manager.getFactory("ods.odfdom.OdsOdfdomDocumentFactory");
		factory = manager.getFactory("ods.jopendocument12.OdsJOpenDocumentFactory");
		
		factory = manager.getFactory("xls.jxl.XlsJxlDocumentFactory");
		factory = manager.getFactory("xls.poi.XlsPoiDocumentFactory");
		
	}

}
