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
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.dom.element.table.TableTableColumnElement;
import org.odftoolkit.odfdom.dom.element.table.TableTableElement;
import org.odftoolkit.odfdom.dom.style.OdfStyleFamily;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetWriterCursorImpl;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/
/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

/**
 * The value writer for Apache odfdom.
 *
 */
class OdsOdfdomDocumentWriter extends AbstractSpreadsheetDocumentWriter
		implements SpreadsheetDocumentWriter {
	/** delegation value with definition of createNew */
	private final class OdsOdfdomDocumentWriterTrait extends
			AbstractOdsOdfdomDocumentTrait<SpreadsheetWriter> {
		OdsOdfdomDocumentWriterTrait(final OdfSpreadsheetDocument document) {
			super(document);
		}

		/** {@inheritDoc} */
		@Override
		protected SpreadsheetWriter createNew(
				/*>>> @UnknownInitialization OdsOdfdomDocumentWriterTrait this, */final OdfTable table) {
			final TableTableElement tte = table.getOdfElement();
			final NodeList colsList = tte
					.getElementsByTagName("table:table-column");
			TableTableColumnElement column = (TableTableColumnElement) colsList.item(0);
			while (colsList.getLength() > 1) {
				final Node item = colsList.item(1);
				tte.removeChild(item);
			}
			column.setTableNumberColumnsRepeatedAttribute(1);
			final NodeList rowList = tte
					.getElementsByTagName("table:table-row");
			while (rowList.getLength() > 0) {
				final Node item = rowList.item(0);
				tte.removeChild(item);
			}
			final NodeList rowListAfter = tte
					.getElementsByTagName("table:table-row");
			final int lengthAfter = rowListAfter.getLength();
			assert lengthAfter == 0;
			return new OdsOdfdomWriter(table);
		}
	}

	/** the *internal* workbook */
	private final OdfSpreadsheetDocument document;

	private final OdfOfficeStyles documentStyles;
	/** delegation value */
	private final AbstractOdsOdfdomDocumentTrait<SpreadsheetWriter> documentTrait;
	/** delegation reader */
	private final OdsOdfdomDocumentReader reader;

	private final OdsOdfdomStyleUtility styleUtility;

	/** the logger */
	final Logger logger;

	/**
	 * @param logger
	 *            the logger
	 * @param value
	 *            the *internal* workbook
	 * @param outputStream
	 *            where to write
	 * @throws SpreadsheetException
	 *             if can't open value writer
	 */
	public OdsOdfdomDocumentWriter(final Logger logger,
			final OdsOdfdomStyleUtility styleUtility,
			final OdfSpreadsheetDocument document,
			final/*@Nullable*/OutputStream outputStream)
			throws SpreadsheetException {
		super(logger, outputStream);
		this.styleUtility = styleUtility;
		this.reader = new OdsOdfdomDocumentReader(styleUtility, document);
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
	@Deprecated
	public boolean createStyle(final String styleName, final String styleString) {
		final OdfStyle newStyle = this.documentStyles.newStyle(styleName,
				OdfStyleFamily.TableCell);
		newStyle.setProperties(this.styleUtility.getProperties(styleString));
		return true;
	}

	/** {@inheritDoc} */
	@Override
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
		return this.reader.getStyleString(styleName);
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
	public boolean setStyle(final String styleName,
			final WrapperCellStyle wrapperCellStyle) {
		final OdfStyle newStyle = this.documentStyles.newStyle(styleName,
				OdfStyleFamily.TableCell);
		newStyle.setProperties(this.styleUtility
				.getProperties(wrapperCellStyle));
		return true;
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public boolean updateStyle(final String styleName, final String styleString) {
		final OdfStyle existingStyle = this.documentStyles.getStyle(styleName,
				OdfStyleFamily.TableCell);
		existingStyle.setProperties(this.styleUtility
				.getProperties(styleString));
		return true;
	}
}
