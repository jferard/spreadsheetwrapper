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

/**
 * Utility class for 1.2 version.
 * @author Julien
 *
 */
final class OdsJopenDocument1_2Util {
	private OdsJopenDocument1_2Util() {} 
	
	/**
	 * @param cell
	 *           the cell
	 * @return the formula
	 */
	public static String getFormula(final MutableCell<SpreadSheet> cell) {
		final Element element = cell.getElement();
		return element.getAttributeValue(OdsConstants.FORMULA_ATTR_NAME,
				element.getNamespace());
	}

	/**
	 * @param odPackage
	 *           the zip file representation
	 * @return
	 *           a new spreadsheet
	 */
	public static SpreadSheet getSpreadSheet(final ODPackage odPackage) {
		return SpreadSheet.create(odPackage);
	}

	/**
	 * @param spreadSheet the *internal* document.xml
	 * @return the styles as *internal* document.xml
	 */
	public static ODXMLDocument getStyles(final SpreadSheet spreadSheet) {
		final ODXMLDocument res;
		final ODPackage odPackage = spreadSheet.getPackage();
		if (odPackage.isSingle())
			res = odPackage.getContent();
		else
			res = odPackage.getXMLFile("styles.xml");
		return res;
	}

	/**
	 * @return a new *internal* document.xml
	 * @throws SpreadsheetException
	 */
	public static SpreadSheet newSpreadsheetDocument()
			throws SpreadsheetException {
		try {
			final SpreadSheet spreadSheet = SpreadSheet
					.createEmpty(new DefaultTableModel());
			spreadSheet.getPackage().putFile(
					"styles.xml",
					OdsJopenDocument1_2Util.createDocument("office",
							"document-styles", "styles.xml"));
			return spreadSheet;
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}

	/**
	 * @param cell the internal cell
	 * @param type the internal type
	 * @param value the value to set
	 */
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

	/**
	 * @param nsPrefix
	 * @param name name of the document
	 * @param zipEntry ???
	 * @return the *internal* odf document
	 */
	private static Document createDocument(final String nsPrefix,
			final String name, final String zipEntry) {
		final XMLVersion version = XMLVersion.OD;
		final Element root = new Element(name, version.getNS(nsPrefix));
		for (final Namespace nameSpace : version.getALL())
			root.addNamespaceDeclaration(nameSpace);

		return new Document(root);
	}
}
