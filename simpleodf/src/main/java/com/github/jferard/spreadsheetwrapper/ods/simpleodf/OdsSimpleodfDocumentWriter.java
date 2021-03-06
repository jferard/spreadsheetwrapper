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
package com.github.jferard.spreadsheetwrapper.ods.simpleodf;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.ListIterator;
import java.util.NoSuchElementException;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.odftoolkit.odfdom.dom.element.table.TableTableColumnElement;
import org.odftoolkit.odfdom.dom.element.table.TableTableElement;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.simple.table.Table;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.OptionalOutput;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>>
import org.checkerframework.checker.initialization.qual.UnknownInitialization;
import org.checkerframework.checker.nullness.qual.RequiresNonNull;
 */

/**
 *
 */
public class OdsSimpleodfDocumentWriter extends
AbstractSpreadsheetDocumentWriter implements SpreadsheetDocumentWriter {
		private static void cleanEmptyTable(final TableTableElement tableElement) {
			final NodeList colsList = tableElement
					.getElementsByTagName("table:table-column");
			assert colsList.getLength() == 1;
			final TableTableColumnElement column = (TableTableColumnElement) colsList
					.item(0);
			column.setTableNumberColumnsRepeatedAttribute(1);
			final NodeList rowList = tableElement
					.getElementsByTagName("table:table-row");
			while (rowList.getLength() > 1) {
				final Node item = rowList.item(1);
				tableElement.removeChild(item);
			}
			final NodeList rowListAfter = tableElement
					.getElementsByTagName("table:table-row");
			final int lengthAfter = rowListAfter.getLength();
			assert lengthAfter == 1;
		}

	/** internal styles */
	private final OdfOfficeStyles documentStyles;

	/** logger */
	private final Logger logger;

	/** *internal* workbook */
	private final InitializableDocument initializableDocument;

	/** the style helper */
	private final OdsSimpleodfStyleHelper styleHelper;

	/**
	 * @param logger
	 *            the logger
	 * @param value
	 *            *internal* workbook
	 * @param outputURL
	 *            URL where to write, should be a local file
	 * @throws SpreadsheetException
	 *             if the value writer can't be created
	 */
	public OdsSimpleodfDocumentWriter(final Logger logger,
			final OdsSimpleodfStyleHelper styleHelper,
			final InitializableDocument initializableDocument, final OptionalOutput optionalOutput)
					throws SpreadsheetException {
		super(logger, optionalOutput);
		this.styleHelper = styleHelper;
		this.logger = logger;
		this.initializableDocument = initializableDocument;
		this.documentStyles = this.initializableDocument.getStyles();
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		try {
			this.optionalOutput.close();
		} catch (final IOException e) {
			final String message = e.getMessage();
			this.logger.log(Level.SEVERE, message == null ? "" : message, e);
		}
		this.initializableDocument.close();
	}

	/** {@inheritDoc} */
	@Override
	public WrapperCellStyle getCellStyle(final String styleName) {
		return this.styleHelper.getCellStyle(this.documentStyles, styleName);
	}

	/** {@inheritDoc} */
	@Override
	public int getSheetCount() {
		return this.getTableList().size();
	}
	
	
	/** {@inheritDoc} */
	@Override
	public List<String> getSheetNames() {
		final List<Table> tables = this.getTableList();
		final List<String> sheetNames = new ArrayList<String>(tables.size());
		for (final Table table : tables)
			sheetNames.add(table.getTableName());
		return sheetNames;
	}

	/**
	 * @param index
	 *            index of the spreadsheet in the document
	 * @return the spreadsheet reader or writer
	 */
	@Override
	public SpreadsheetWriter getSpreadsheet(final int index) {
		final SpreadsheetWriter spreadsheet;
		if (this.accessor.hasByIndex(index))
			spreadsheet = this.accessor.getByIndex(index);
		else {
			final List<Table> tables = this.getTableList();
			if (index < 0 || index >= tables.size())
				throw new IndexOutOfBoundsException(String.format(
						"No sheet at position %d", index));

			final Table table = tables.get(index);
			spreadsheet = this.createNew(table);
			this.accessor.put(table.getTableName(), index, spreadsheet);
		}
		return spreadsheet;
	}

	/**
	 * @return a list of internal tables
	 */
	/*>>> @RequiresNonNull("initializableDocument")*/
	public final List<Table> getTableList(/*>>> @UnknownInitialization AbstractOdsSimpleodfDocumentDelegate<T> this*/) {
		return this.initializableDocument.getTableList();
	}

	/** {@inheritDoc} */
	@Override
	public void save() throws SpreadsheetException {
		OutputStream outputStream = null;
		try {
			outputStream = this.optionalOutput.getStream();
			if (outputStream == null)
				throw new IllegalStateException(
						String.format("Use saveAs when optionalOutput file is not specified"));
			this.initializableDocument.save(outputStream);
		} catch (final Exception e) { // NOPMD by Julien on 02/09/15 22:55
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
		return this.styleHelper.setStyle(this.documentStyles, styleName,
				wrapperCellStyle);
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetWriter addSheetWithCheckedIndex(final int index, final String sheetName)
			throws CantInsertElementInSpreadsheetException {
		Table table = this.initializableDocument.addTable(index, sheetName);
		final TableTableElement tableElement = table.getOdfElement();
		cleanEmptyTable(tableElement);

		final SpreadsheetWriter spreadsheet = this.createNew(table);
		this.accessor.put(sheetName, index, spreadsheet);
		return spreadsheet;
	}

	/*@RequiresNonNull("delegateStyleHelper")*/
	protected SpreadsheetWriter createNew(
			/*>>> @UnknownInitialization OdsSimpleodfDocumentWriterDelegate this, */final Table table) {
		return new OdsSimpleodfWriter(table, this.styleHelper);
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetWriter findSpreadsheetAndCreateReaderOrWriter(final String sheetName) {
		final SpreadsheetWriter spreadsheet;
		final List<Table> tables = this.getTableList();
		final ListIterator<Table> tablesIterator = tables.listIterator();
		Table table;
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
	
}
