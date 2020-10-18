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
package com.github.jferard.spreadsheetwrapper.examples;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;
import java.util.Random;
import java.util.logging.Logger;

import javax.swing.table.DefaultTableModel;

import org.jdom.Document;
import org.jdom.Element;
import org.jdom.Namespace;
import org.jopendocument.dom.XMLVersion;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;
import org.junit.Test;

import com.github.jferard.spreadsheetwrapper.DocumentFactoryManager;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;

/**
 * Unit test for Examples.
 */
public class BenchTest {
	private static final int ROWS = 50000;

	private static SpreadsheetDocumentWriter createSimpleDocument(String lib)
			throws SpreadsheetException {
		final DocumentFactoryManager manager = new DocumentFactoryManager(
				Logger.getAnonymousLogger());
		final SpreadsheetDocumentFactory factory = manager.getFactory(lib);
		return factory.create(new File(lib+"temp.ods"));
	}

//	@Test
	public void OdftoolkitBenchTest() throws SpreadsheetException {
		this.benchTest("ods.odfdom.OdsOdfdomDocumentFactory");
	}
	
//	@Test
	public void jOpendocument12BenchTest() throws SpreadsheetException {
		this.benchTest("ods.jopendocument1_2.OdsJOpenDocumentFactory");
	}
	
	@Test
	public void SimpleOdsBenchTest() throws SpreadsheetException {
		this.benchTest("ods.simpleods.OdsSimpleodsDocumentFactory");
	}	
	
//	@Test
	public void jOpendocument12BASICBenchTest() throws SpreadsheetException {
        SpreadSheet spreadSheet = newJSpreadsheetDocument();
        System.out.println("Filling a 25000 rows, 20 columns spreadsheet");
        long t1 = System.currentTimeMillis();

        final Sheet sheet = spreadSheet.getSheet(0);
        
        final Random random = new Random();
        for (int x = 1; x < 21; x++) {
            for (int y = 1; y < ROWS+1; y++) {
        		sheet.ensureColumnCount(x+1);
        		sheet.ensureRowCount(y+1);
                sheet.setValueAt(random.nextInt(1000), x, y);
            }
        }

        long t2 = System.currentTimeMillis();
        System.out.println("Filled in " + (t2 - t1) + " ms");
	}
	

	public void benchTest(String lib) throws SpreadsheetException {
		SpreadsheetDocumentWriter documentWriter = createSimpleDocument(lib);
		System.out.println("Filling a 25000 rows, 20 columns spreadsheet :"+lib);
		long t1 = System.currentTimeMillis();
		for (String name : Arrays.asList("fristsheet", "secondsheet", "thirdsheet")) {
			SpreadsheetWriter w =	documentWriter.addSheet(name);
			
	
			final Random random = new Random();
			for (int r = 0; r < ROWS; r++) {
				for (int c = 0; c < 20; c++) {
					w.setCellContent(r, c, random.nextInt(1000));
				}
			}
	
		}
		long t2 = System.currentTimeMillis();
		System.out.println("Filled in " + (t2 - t1) + " ms");
		documentWriter.save();
	}
	
	/**
	 * @return a new *internal* document.xml
	 * @throws SpreadsheetException
	 */
	public static SpreadSheet newJSpreadsheetDocument()
			throws SpreadsheetException {
		try {
			final SpreadSheet spreadSheet = SpreadSheet
					.createEmpty(new DefaultTableModel());
			spreadSheet.getPackage().putFile(
					"styles.xml",
					createJDocument("office",
							"document-styles", "styles.xml"));
			return spreadSheet;
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}
	
	/**
	 * @param nsPrefix
	 * @param name name of the document
	 * @param zipEntry ???
	 * @return the *internal* odf document
	 */
	private static Document createJDocument(final String nsPrefix,
			final String name, final String zipEntry) {
		final XMLVersion version = XMLVersion.OD;
		final Element root = new Element(name, version.getNS(nsPrefix));
		for (final Namespace nameSpace : version.getALL())
			root.addNamespaceDeclaration(nameSpace);

		return new Document(root);
	}
}
