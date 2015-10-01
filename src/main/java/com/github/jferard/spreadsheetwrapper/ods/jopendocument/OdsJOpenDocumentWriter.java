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

package com.github.jferard.spreadsheetwrapper.ods.jopendocument;

import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.jdom.Attribute;
import org.jdom.Element;
import org.jdom.Namespace;
import org.jopendocument.dom.ODXMLDocument;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetWriterCursorImpl;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/
/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

/**
 * The spreadSheet writer for Apache JOpen.
 *
 */
class OdsJOpenDocumentWriter extends AbstractSpreadsheetDocumentWriter
		implements SpreadsheetDocumentWriter {
	/** delegation spreadSheet with definition of createNew */
	private final class OdsJOpenDocumentWriterTrait extends
			AbstractOdsJOpenDocumentTrait<SpreadsheetWriter> {
		OdsJOpenDocumentWriterTrait(final SpreadSheet document) {
			super(document);
		}

		/** {@inheritDoc} */
		@Override
		protected SpreadsheetWriter createNew(
				/*>>> @UnknownInitialization OdsJOpenDocumentWriterTrait this, */final Sheet sheeet) {
			return new OdsJOpenWriter(sheeet);
		}
	}

	/** the *internal* workbook */
	private final SpreadSheet spreadSheet;

	/** delegation spreadSheet */
	private final AbstractOdsJOpenDocumentTrait<SpreadsheetWriter> documentTrait;
	/** delegation reader */
	private final OdsJOpenDocumentReader reader;
	/** the logger */
	final Logger logger;

	private HashMap<String, String> styles;

	/**
	 * @param logger
	 *            the logger
	 * @param spreadSheet
	 *            the *internal* workbook
	 * @param outputStream
	 *            where to write
	 * @throws SpreadsheetException
	 *             if can't open spreadSheet writer
	 */
	public OdsJOpenDocumentWriter(final Logger logger,
			final SpreadSheet document,
			final/*@Nullable*/OutputStream outputStream)
			throws SpreadsheetException {
		super(logger, outputStream);
		this.reader = new OdsJOpenDocumentReader(document);
		this.logger = logger;
		this.spreadSheet = document;
		this.documentTrait = new OdsJOpenDocumentWriterTrait(document);
		this.styles = new HashMap<String, String>();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter addSheet(final int index, final String sheetName)
			throws IndexOutOfBoundsException,
			CantInsertElementInSpreadsheetException {
		return this.documentTrait.addSheet(index, sheetName);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter addSheet(final String sheetName)
			throws CantInsertElementInSpreadsheetException {
		return this.documentTrait.addSheet(sheetName);
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		this.reader.close();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriterCursor getNewCursorByIndex(final int index)
			throws SpreadsheetException {
		return new SpreadsheetWriterCursorImpl(this.getSpreadsheet(index));
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriterCursor getNewCursorByName(final String sheetName)
			throws SpreadsheetException {
		return new SpreadsheetWriterCursorImpl(this.getSpreadsheet(sheetName));
	}

	/** {@inheritDoc} */
	@Override
	public int getSheetCount() {
		return this.reader.getSheetCount();
	}

	/** {@inheritDoc} */
	@Override
	public List<String> getSheetNames() {
		return this.reader.getSheetNames();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter getSpreadsheet(final int index) {
		return this.documentTrait.getSpreadsheet(index);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter getSpreadsheet(final String sheetName) {
		return this.documentTrait.getSpreadsheet(sheetName);
	}

	/** */
	@Override
	public void save() throws SpreadsheetException {
		if (this.outputStream == null)
			throw new IllegalStateException(
					String.format("Use saveAs when output file is not specified"));
		try {
			this.spreadSheet.getPackage().save(this.outputStream);
		} catch (final Exception e) { // NOPMD by Julien on 03/09/15 22:09
			this.logger.log(Level.SEVERE, String.format(
					"this.spreadsheetDocument.save(%s) not ok",
					this.outputStream), e);
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public boolean createStyle(String styleName, String styleString) {
		ODXMLDocument stylesDoc = this.spreadSheet.getPackage().getStyles();
		Element namedStyles = stylesDoc.getChild("styles", true);
		List<Element> styles = namedStyles.getChildren("style");

		for (Element style : styles) {
			if (styleName.equals(style.getAttributeValue("name")))
				return false;
		}

		Map<String, String> propertiesMap = this.getPropertiesMap(styleString);
		Element style = createStyle(styleName, propertiesMap);
		namedStyles.addContent(style);
		return true;
	}

	private Element createStyle(String styleName, Map<String, String> propertiesMap) {
		Namespace officeNS = Namespace.getNamespace("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
		Namespace foNS = Namespace.getNamespace("fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0");
		Namespace styleNS = Namespace.getNamespace("style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0");
		Element style = new Element("style", styleNS);
		style.setAttribute("name", styleName, styleNS);
		style.setAttribute("family", "table-cell", styleNS);
		if (propertiesMap.containsKey("background-color")) {
			Element tableCellProps = new Element("table-cell-properties", styleNS);
			tableCellProps.setAttribute("background-color",
					propertiesMap.get("background-color"), foNS);
			style.addContent(tableCellProps);
		}
		if (propertiesMap.containsKey("font-weight")) {
			Element textProps = new Element("text-properties", styleNS);
			textProps.setAttribute("font-weight",
					propertiesMap.get("font-weight"), foNS);
			style.addContent(textProps);
		}
		return style;
	}

	/** {@inheritDoc} */
	@Override
	public boolean updateStyle(String styleName, String styleString) {
		ODXMLDocument stylesDoc = this.spreadSheet.getPackage().getStyles();
		Element namedStyles = stylesDoc.getChild("styles", true);
		List<Element> styles = namedStyles.getChildren("style");

		for (Element style : styles) {
			if (styleName.equals(style.getAttributeValue("name"))) {
				styles.remove(style);
				Map<String, String> propertiesMap = this.getPropertiesMap(styleString);
				Element style2 = createStyle(styleName, propertiesMap);
				namedStyles.addContent(style);
				return true;
			}
		}
		return false;
	}

	/** {@inheritDoc} */
	@Override
	public String getStyleString(String styleName) {
		throw new UnsupportedOperationException();
	}
}
