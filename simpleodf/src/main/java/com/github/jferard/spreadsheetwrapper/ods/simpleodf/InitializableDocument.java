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

import java.io.OutputStream;
import java.util.Collections;
import java.util.List;
import java.util.Locale;

import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.table.Table;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;

/**
 * A InitializableDocument is a kinf of Proxy for SpreadsheetDocument. Wraps all
 * the management of initialization (= handling the dummy table).
 */
class InitializableDocument {

	static final InitializableDocument createInitialized(
			SpreadsheetDocument document) {
		return new InitializableDocument(document, true);
	}

	static final InitializableDocument createUninitialized(
			SpreadsheetDocument document) {
		return new InitializableDocument(document, false);
	}

	/** true if initialized */
	private boolean initialized;

	/** the object */
	private final SpreadsheetDocument document;
	
	/**
	 * @param document
	 *            a stateful SpreadsheetDocument, ie a SpreadsheetDocument that
	 *            is, or not, initialized
	 */
	private InitializableDocument(SpreadsheetDocument document,
			boolean initialized) {
		this.document = document;
		this.initialized = initialized;
	}

	/** Adds a table at an index */
	public Table addTable(final int index, final String sheetName)
			throws CantInsertElementInSpreadsheetException {
		Table tempTable = this.document.getTableByName(sheetName);
		if (tempTable != null)
			throw new IllegalArgumentException(
					String.format("Sheet %s exists", sheetName));

		Table table;
		final List<Table> tables = this.document.getTableList();
		final int tablesCount = tables.size();
		if (this.initialized) {
			if (index == tablesCount)
				table = Table.newTable(this.document);
			else
				table = this.document.insertSheet(index);
			if (table == null)
				throw new CantInsertElementInSpreadsheetException();
		} else {
			assert tablesCount == 1 && index == 0;
			table = tables.get(0);
			table.setTableName(sheetName);
			this.initialized = true;
		}
		return table;
	}

	/**
	 * close the document
	 */
	public void close() {
		this.document.close();
	}

	/**
	 * @return the document with the odf styles
	 */
	public OdfOfficeStyles getStyles() {
		return this.document.getDocumentStyles();
	}

	/**
	 * @return the list of the sheets in the document
	 */
	public List<Table> getTableList() {
		final List<Table> tables;
		if (this.initialized)
			tables = this.document.getTableList();
		else
			tables = Collections.emptyList();
		return tables;
	}

	/**
	 * @param outputStream
	 *            the stream to write to
	 * @throws Exception
	 *             if odftoolkit throws an exception !
	 */
	public void save(final OutputStream outputStream) throws Exception {
		this.document.save(outputStream);
	}

	/**
	 * @param locale
	 *            the locale to set
	 */
	public void setLocale(final Locale locale) {
		this.document.setLocale(locale);
	}

}
