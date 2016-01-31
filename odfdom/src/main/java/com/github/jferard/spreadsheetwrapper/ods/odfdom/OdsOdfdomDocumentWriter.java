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

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.ListIterator;
import java.util.NoSuchElementException;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.dom.element.table.TableTableColumnElement;
import org.odftoolkit.odfdom.dom.element.table.TableTableElement;
import org.odftoolkit.odfdom.dom.style.OdfStyleFamily;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.github.jferard.spreadsheetwrapper.Output;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetWriterCursorImpl;
import com.github.jferard.spreadsheetwrapper.ods.apache.OdsOdfdomStyleHelper;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>>
import org.checkerframework.checker.initialization.qual.UnknownInitialization;
import org.checkerframework.checker.nullness.qual.RequiresNonNull;
 */

/**
 * The value writer for Apache odfdom.
 *
 */
class OdsOdfdomDocumentWriter extends AbstractSpreadsheetDocumentWriter
implements SpreadsheetDocumentWriter {
	/** the *internal* workbook */
	private final OdfSpreadsheetDocument document;

	/** the document styles */
	private final OdfOfficeStyles documentStyles;

	/** the logger */
	private final Logger logger;

	/** helper object for style */
	private final OdsOdfdomStyleHelper styleHelper;

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
	OdsOdfdomDocumentWriter(final Logger logger,
			final OdsOdfdomStyleHelper styleHelper,
			final OdfSpreadsheetDocument document, final Output output)
			throws SpreadsheetException {
		super(logger, output);
		this.styleHelper = styleHelper;
		this.logger = logger;
		this.document = document;
		this.documentStyles = this.document.getDocumentStyles();
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		try {
			this.output.close();
		} catch (final IOException e) {
			final String message = e.getMessage();
			this.logger.log(Level.SEVERE, message == null ? "" : message, e);
		}
		this.document.close();
	}

	/** {@inheritDoc} */
	@Override
	public WrapperCellStyle getCellStyle(final String styleName) {
		final OdfStyle existingStyle = this.documentStyles.getStyle(styleName,
				OdfStyleFamily.TableCell);
		return this.styleHelper.toWrapperCellStyle(existingStyle);
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
		final List<OdfTable> tables = this.document.getTableList();
		return tables.size();
	}

	/** {@inheritDoc} */
	@Override
	public List<String> getSheetNames() {
		final List<OdfTable> tables = this.document.getTableList();
		final List<String> tableNames = new ArrayList<String>(tables.size());
		for (final OdfTable table : tables)
			tableNames.add(table.getTableName());
		return tableNames;
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter getSpreadsheet(final int index) {
		final SpreadsheetWriter spreadsheet;
		if (this.accessor.hasByIndex(index))
			spreadsheet = this.accessor.getByIndex(index);
		else {
			final List<OdfTable> tables = this.document.getTableList();
			if (index < 0 || index >= tables.size())
				throw new IndexOutOfBoundsException(String.format(
						"No sheet at position %d", index));

			final OdfTable table = this.document.getTableList().get(index);
			spreadsheet = this.createNew(table);
			this.accessor.put(table.getTableName(), index, spreadsheet);
		}
		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	public void save() throws SpreadsheetException {
		OutputStream outputStream = null;
		try {
			outputStream = this.output.getStream();
			if (outputStream == null)
				throw new IllegalStateException(
						String.format("Use saveAs when output file/stream is not specified"));
			this.document.save(outputStream);
		} catch (final Exception e) { // NOPMD by Julien on 03/09/15 22:09
			this.logger.log(Level.SEVERE, String.format(
					"this.spreadsheetDocument.save(%s) not ok", outputStream),
					e);
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyle(final String styleName,
			final WrapperCellStyle wrapperCellStyle) {
		final OdfStyle newStyle = this.documentStyles.newStyle(styleName,
				OdfStyleFamily.TableCell);
		this.styleHelper.setWrapperCellStyle(newStyle, wrapperCellStyle);
		return true;
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetWriter addSheetWithCheckedIndex(final int index, final String sheetName) {
		if (index != this.getSheetCount())
			throw new UnsupportedOperationException(String.format(
					"Should insert sheet at index %d only",
					this.getSheetCount()));

		OdfTable table = this.document.getTableByName(sheetName);
		if (table != null)
			throw new IllegalArgumentException(String.format("Sheet %s exists",
					sheetName));

		table = OdfTable.newTable(this.document);
		final TableTableElement tableElement = table.getOdfElement();
		cleanEmptyTable(tableElement);
		tableElement.setTableNameAttribute(sheetName);
		final SpreadsheetWriter spreadsheet = this.createNew(table);
		// the table is added at the end
		this.accessor.put(sheetName, index, spreadsheet);
		return spreadsheet;
	}

	/**
	 * Create a new reader/writer
	 *
	 * @param table
	 *            *internal* table
	 * @return the reader/writer
	 */
	/**
	 * @param table
	 * @return
	 */
	/*@RequiresNonNull("delegateStyleHelper")*/
	protected SpreadsheetWriter createNew(final OdfTable table) {
		return new OdsOdfdomWriter(table, this.styleHelper);
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetWriter findSpreadsheetAndCreateReaderOrWriter(final String sheetName) {
		final SpreadsheetWriter spreadsheet;
		final List<OdfTable> tableList = this.document.getTableList();
		final ListIterator<OdfTable> tablesIterator = tableList.listIterator();
		OdfTable table;
		while (tablesIterator.hasNext()) {
			final int index = tablesIterator.nextIndex();
			table = tablesIterator.next();
			if (table.getTableName().equals(sheetName)) {
				spreadsheet = this.createNew(table);
				this.accessor.put(sheetName, index, spreadsheet);
				return spreadsheet;
			}
		}
		throw new NoSuchElementException(String.format(
				"No %s sheet in workbook", sheetName));
	}
	
	/**
	 * Cleans the table : remove everything that odftoolkit adds, but let
	 * <table:table>
	 * <table:table-column number-columns-repeated="1" />
	 * </table:table>
	 *
	 * @param tableElement
	 *            the odf element
	 */
	private static void cleanEmptyTable(final TableTableElement tableElement) {
		final NodeList colsList = tableElement
				.getElementsByTagName("table:table-column");
		final TableTableColumnElement column = (TableTableColumnElement) colsList
				.item(0);
		while (colsList.getLength() > 1) {
			final Node item = colsList.item(1);
			tableElement.removeChild(item);
		}
		column.setTableNumberColumnsRepeatedAttribute(1);
		final NodeList rowList = tableElement
				.getElementsByTagName("table:table-row");
		while (rowList.getLength() > 0) {
			final Node item = rowList.item(0);
			tableElement.removeChild(item);
		}
		final NodeList rowListAfter = tableElement
				.getElementsByTagName("table:table-row");
		assert rowListAfter.getLength() == 0;
	}
	
}
