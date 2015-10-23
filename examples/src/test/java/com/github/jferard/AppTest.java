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
