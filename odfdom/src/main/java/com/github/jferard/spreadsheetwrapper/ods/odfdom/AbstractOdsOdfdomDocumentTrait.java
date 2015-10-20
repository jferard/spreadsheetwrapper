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

import java.util.List;
import java.util.ListIterator;
import java.util.NoSuchElementException;

import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.dom.element.table.TableTableElement;

import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentTrait;

/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

abstract class AbstractOdsOdfdomDocumentTrait<T> extends
		AbstractSpreadsheetDocumentTrait<T> {
	/** the *internal* value (workbook) */
	private final OdfSpreadsheetDocument document;

	/**
	 * @param value
	 *            the *internal* value (workbook)
	 */
	public AbstractOdsOdfdomDocumentTrait(final OdfSpreadsheetDocument document) {
		super();
		this.document = document;
		final List<OdfTable> tables = this.document.getTableList();
		final ListIterator<OdfTable> tablesIterator = tables.listIterator();
		while (tablesIterator.hasNext()) {
			final Integer index = tablesIterator.nextIndex();
			final OdfTable table = tablesIterator.next();
			final String name = table.getTableName();
			final T reader = this.createNew(table);
			this.accessor.put(name, index, reader);
		}
	}

	/**
	 * @param index
	 *            index of the sheet
	 * @return the reader/writer
	 */
	public T getSpreadsheet(final int index) {
		final T spreadsheet;
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
	protected T addSheetWithCheckedIndex(final int index, final String sheetName) {
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
		tableElement.setTableNameAttribute(sheetName);
		final T spreadsheet = this.createNew(table);
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
	protected abstract T createNew(
			/*>>> @UnknownInitialization AbstractOdsOdfdomDocumentTrait<T> this, */final OdfTable table);

	/** {@inheritDoc} */
	@Override
	protected T findSpreadsheetAndCreateReaderOrWriter(final String sheetName) {
		final T spreadsheet;
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

	/** {@inheritDoc} */
	@Override
	protected int getSheetCount() {
		return this.document.getTableList().size();
	}
}