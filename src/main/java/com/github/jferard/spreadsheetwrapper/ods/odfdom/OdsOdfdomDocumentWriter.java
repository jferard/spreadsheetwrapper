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

package com.github.jferard.spreadsheetwrapper.ods.odfdom;

import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.dom.style.OdfStyleFamily;
import org.odftoolkit.odfdom.dom.style.props.OdfStyleProperty;
import org.odftoolkit.odfdom.dom.style.props.OdfTableCellProperties;
import org.odftoolkit.odfdom.dom.style.props.OdfTextProperties;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;

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
 * The document writer for Apache odfdom.
 *
 */
class OdsOdfdomDocumentWriter extends AbstractSpreadsheetDocumentWriter
		implements SpreadsheetDocumentWriter {
	/** delegation document with definition of createNew */
	private final class OdsOdfdomDocumentWriterTrait extends
			AbstractOdsOdfdomDocumentTrait<SpreadsheetWriter> {
		OdsOdfdomDocumentWriterTrait(final OdfSpreadsheetDocument document) {
			super(document);
		}

		/** {@inheritDoc} */
		@Override
		protected SpreadsheetWriter createNew(
				/*>>> @UnknownInitialization OdsOdfdomDocumentWriterTrait this, */final OdfTable table) {
			return new OdsOdfdomWriter(table);
		}
	}

	/** the *internal* workbook */
	private final OdfSpreadsheetDocument document;

	/** delegation document */
	private final AbstractOdsOdfdomDocumentTrait<SpreadsheetWriter> documentTrait;
	/** delegation reader */
	private final OdsOdfdomDocumentReader reader;
	/** the logger */
	final Logger logger;

	private OdfOfficeStyles documentStyles;

	/**
	 * @param logger
	 *            the logger
	 * @param document
	 *            the *internal* workbook
	 * @param outputStream
	 *            where to write
	 * @throws SpreadsheetException
	 *             if can't open document writer
	 */
	public OdsOdfdomDocumentWriter(final Logger logger,
			final OdfSpreadsheetDocument document,
			final/*@Nullable*/OutputStream outputStream)
			throws SpreadsheetException {
		super(logger, outputStream);
		this.reader = new OdsOdfdomDocumentReader(document);
		this.logger = logger;
		this.document = document;
		this.documentStyles = this.document.getDocumentStyles();
		this.documentTrait = new OdsOdfdomDocumentWriterTrait(document);
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
			this.document.save(this.outputStream);
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
		OdfStyle newStyle = this.documentStyles.newStyle(styleName, OdfStyleFamily.TableCell);
		newStyle.setProperties(this.getProperties(styleString));
		return true;
	}
	
	private Map<OdfStyleProperty, String> getProperties(String styleString) {
		Map<OdfStyleProperty, String> properties = new HashMap<OdfStyleProperty, String>();
		String[] styleProps = styleString.split(";");
		String[] entry;
		for (String styleProp : styleProps) { 
			entry = styleProp.split(":");
			if (entry[0].equals("font-weight")) {
				properties.put(OdfTextProperties.FontWeight, entry[1]);
			} else if (entry[0].equals("background-color")) {
				properties.put(OdfTableCellProperties.BackgroundColor, entry[1]);
			}
		}
		return properties;
	}

	/** {@inheritDoc} */
	@Override
	public boolean updateStyle(String styleName, String styleString) {
		OdfStyle existingStyle = this.documentStyles.getStyle(styleName, OdfStyleFamily.TableCell);
		existingStyle.setProperties(this.getProperties(styleString));
		return true;
	}
	
	/** {@inheritDoc} */
	@Override
	public String getStyleString(String styleName) {
		OdfStyle existingStyle = this.documentStyles.getStyle(styleName, OdfStyleFamily.TableCell);
		return this.getStyleString(existingStyle);
	}

	private String getStyleString(OdfStyle style) {
		String fontWeight = style.getProperty(OdfTextProperties.FontWeight);
		String backgroundColor = style.getProperty(OdfTableCellProperties.BackgroundColor);
		return String.format("font-weight:%s;background-color:%s", fontWeight, backgroundColor);  
	}	
}
