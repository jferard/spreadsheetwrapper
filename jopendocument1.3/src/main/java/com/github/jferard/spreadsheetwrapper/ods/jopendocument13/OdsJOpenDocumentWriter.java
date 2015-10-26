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

package com.github.jferard.spreadsheetwrapper.ods.jopendocument13;

import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.jdom.Element;
import org.jopendocument.dom.ODXMLDocument;
import org.jopendocument.dom.spreadsheet.Sheet;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.Output;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetWriterCursorImpl;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/
/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

/**
 * The sfSpreadSheet writer for Apache JOpen.
 *
 */
/**
 * @author Julien
 *
 */
class OdsJOpenDocumentWriter extends AbstractSpreadsheetDocumentWriter
implements SpreadsheetDocumentWriter {
	/** delegation sfSpreadSheet with definition of createNew */
	private final class OdsJOpenDocumentWriterTrait extends
	AbstractOdsJOpenDocumentTrait<SpreadsheetWriter> {
		OdsJOpenDocumentWriterTrait(final OdsJOpenStatefulDocument sfSpreadSheet) {
			super(sfSpreadSheet);
		}

		/** {@inheritDoc} */
		@Override
		protected SpreadsheetWriter createNew(
				/*>>> @UnknownInitialization OdsJOpenDocumentWriterTrait this, */final Sheet sheeet) {
			return new OdsJOpenWriter(sheeet);
		}
	}

	/**
	 * XPath could do better...
	 *
	 * @param name
	 *            the name to find in attributes
	 * @param elements
	 *            the elements to look at
	 * @return the matching element, or null
	 */
	private static Element findElementWithName(final String name,
			final List<Element> elements) {
		for (final Element element : elements) {
			if (name.equals(element.getAttributeValue("name")))
				return element;
		}
		return null;
	}

	/** delegation sfSpreadSheet */
	private final AbstractOdsJOpenDocumentTrait<SpreadsheetWriter> documentTrait;
	/** the logger */
	private final Logger logger;
	/** delegation reader */
	private final OdsJOpenDocumentReader reader;

	/** the *internal* workbook */
	private final OdsJOpenStatefulDocument sfSpreadSheet;

	/** utility for style */
	private final OdsJOpenStyleUtility styleUtility;

	/**
	 * @param logger
	 *            the logger
	 * @param sfSpreadSheet
	 *            the *internal* workbook
	 * @param outputStream
	 *            where to write
	 * @param newDocument
	 * @throws SpreadsheetException
	 *             if can't open sfSpreadSheet writer
	 */
	public OdsJOpenDocumentWriter(final Logger logger,
			final OdsJOpenStyleUtility styleUtility,
			final OdsJOpenStatefulDocument sfSpreadSheet, final Output output)
					throws SpreadsheetException {
		super(logger, output);
		this.styleUtility = styleUtility;
		this.reader = new OdsJOpenDocumentReader(sfSpreadSheet);
		this.logger = logger;
		this.sfSpreadSheet = sfSpreadSheet;
		this.documentTrait = new OdsJOpenDocumentWriterTrait(sfSpreadSheet);
		new HashMap<String, String>();
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
		try {
			this.output.close();
		} catch (final IOException e) {
			this.logger.log(Level.SEVERE, e.getMessage(), e);
		}
		this.reader.close();
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public boolean createStyle(final String styleName, final String styleString) {
		final ODXMLDocument stylesDoc = this.sfSpreadSheet.getStyles();
		final Element namedStyles = stylesDoc.getChild("styles", true);
		@SuppressWarnings("unchecked")
		final List<Element> styles = namedStyles.getChildren("style");
		Element oldStyle = OdsJOpenDocumentWriter.findElementWithName(
				styleName, styles);
		if (oldStyle == null) {
			oldStyle = this.styleUtility.createBaseStyle(styleName);
			namedStyles.addContent(oldStyle);
		} else
			oldStyle.removeContent();

		this.styleUtility.setStyle(oldStyle, styleString);
		return true;
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public WrapperCellStyle getCellStyle(final String styleName) {
		return this.reader.getCellStyle(styleName);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriterCursor getNewCursorByIndex(final int index) {
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

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public String getStyleString(final String styleName) {
		throw new UnsupportedOperationException();
	}

	/** */
	@Override
	public void save() throws SpreadsheetException {
		OutputStream outputStream = null;
		try {
			outputStream = this.output.getStream();
			if (outputStream == null)
				throw new IllegalStateException(
						String.format("Use saveAs when output file/stream is not specified"));
			this.sfSpreadSheet.save(outputStream);
		} catch (final Exception e) { // NOPMD by Julien on 03/09/15 22:09
			this.logger.log(Level.SEVERE, String.format(
					"this.spreadsheetDocument.save(%s) not ok", outputStream),
					e);
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	/* (non-Javadoc)
	 * @see com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter#setStyle(java.lang.String, com.github.jferard.spreadsheetwrapper.WrapperCellStyle)
	 */
	@Override
	public boolean setStyle(final String styleName,
			final WrapperCellStyle wrapperCellStyle) {
		final ODXMLDocument stylesDoc = this.sfSpreadSheet.getStyles();
		final Element namedStyles = stylesDoc.getChild("styles", true);
		@SuppressWarnings("unchecked")
		final List<Element> styles = namedStyles.getChildren("style");
		Element newStyle = OdsJOpenDocumentWriter.findElementWithName(
				styleName, styles);
		if (newStyle == null) {
			newStyle = this.styleUtility.createBaseStyle(styleName);
			namedStyles.addContent(newStyle);
		} else
			newStyle.removeContent();

		this.styleUtility.setStyle(newStyle, wrapperCellStyle);
		return true;
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public boolean updateStyle(final String styleName, final String styleString) {
		final ODXMLDocument stylesDoc = this.sfSpreadSheet.getStyles();
		final Element namedStyles = stylesDoc.getChild("styles", true);
		@SuppressWarnings("unchecked")
		final List<Element> styles = namedStyles.getChildren("style");
		Element oldStyle = OdsJOpenDocumentWriter.findElementWithName(
				styleName, styles);
		if (oldStyle == null) {
			oldStyle = this.styleUtility.createBaseStyle(styleName);
			this.styleUtility.setStyle(oldStyle, styleString);
			namedStyles.addContent(oldStyle);
		}
		return oldStyle == null;
	}
}
