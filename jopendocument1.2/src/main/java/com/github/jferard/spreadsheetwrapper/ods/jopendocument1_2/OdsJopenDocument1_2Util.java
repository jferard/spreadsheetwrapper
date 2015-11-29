package com.github.jferard.spreadsheetwrapper.ods.jopendocument1_2;

import java.io.IOException;

import javax.swing.table.DefaultTableModel;

import org.jdom.Document;
import org.jdom.Element;
import org.jdom.Namespace;
import org.jdom.Text;
import org.jopendocument.dom.ODPackage;
import org.jopendocument.dom.ODValueType;
import org.jopendocument.dom.ODXMLDocument;
import org.jopendocument.dom.XMLVersion;
import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;

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
	
	public static String getFormula(final MutableCell<SpreadSheet> cell) {
		final Element element = cell.getElement();
		return element.getAttributeValue(OdsConstants.FORMULA_ATTR_NAME,
				element.getNamespace());
	}

	public static ODXMLDocument getStyles(SpreadSheet spreadSheet) {
		final ODXMLDocument res;
		final ODPackage odPackage = spreadSheet.getPackage();
		if (odPackage.isSingle())
			res = odPackage.getContent();
		else
			res = odPackage.getXMLFile("styles.xml");
		return res;
	}
	
	public static void setValue(final MutableCell<SpreadSheet> cell,
			final ODValueType type, final Object value) {
		final Element odfElement = cell.getElement();
		final Namespace officeNS = XMLVersion.OD.getOFFICE();
		final Namespace textNS = XMLVersion.OD.getTEXT();

		odfElement.setAttribute("value-type", type.getName(), officeNS);
		odfElement.setAttribute(type.getValueAttribute(), type.format(value),
				officeNS);
		final Element child = odfElement.getChild("p", textNS);
		final Element element = child == null ? new Element("p", textNS)
		: child;
		element.setContent(new Text(value.toString()));
		odfElement.setContent(element);
	}
}
