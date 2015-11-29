package com.github.jferard.spreadsheetwrapper.ods.jopendocument1_2;

import java.io.IOException;
import java.io.OutputStream;

import javax.swing.table.DefaultTableModel;

import org.jdom.Document;
import org.jdom.Element;
import org.jdom.Namespace;
import org.jopendocument.dom.ODPackage;
import org.jopendocument.dom.XMLVersion;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetException;

public class OdsJopenDocument1_2Util {
	public static SpreadSheet getSpreadSheet(final ODPackage odPackage) {
		return SpreadSheet.create(odPackage);
	}
	
	public static SpreadSheet newSpreadsheetDocument()
					throws SpreadsheetException {
		try {
			final SpreadSheet spreadSheet = SpreadSheet
					.createEmpty(new DefaultTableModel());
			spreadSheet.getPackage().putFile(
					"styles.xml",
					OdsJopenDocument1_2Util.createDocument("office", "document-styles",
							"styles.xml"));
			return spreadSheet;
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}

	private static Document createDocument(final String nsPrefix,
			final String name, final String zipEntry) {
		final XMLVersion version = XMLVersion.OD;
		final Element root = new Element(name, version.getNS(nsPrefix));
		for (final Namespace nameSpace : version.getALL())
			root.addNamespaceDeclaration(nameSpace);

		return new Document(root);
	}
}
