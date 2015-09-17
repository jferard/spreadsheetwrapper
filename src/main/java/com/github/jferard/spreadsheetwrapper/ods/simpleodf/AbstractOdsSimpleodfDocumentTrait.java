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

import java.util.List;
import java.util.ListIterator;
import java.util.NoSuchElementException;

import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.table.Table;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentTrait;

/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

public abstract class AbstractOdsSimpleodfDocumentTrait<T> extends
AbstractSpreadsheetDocumentTrait<T> {
	/** the document wrapper for delegation */
	private final SpreadsheetDocument document;

	public AbstractOdsSimpleodfDocumentTrait(final SpreadsheetDocument document) {
		super();
		this.document = document;
		final List<Table> tables = this.document.getTableList();
		final ListIterator<Table> tablesIterator = tables.listIterator();
		while (tablesIterator.hasNext()) {
			final Integer index = tablesIterator.nextIndex();
			final Table table = tablesIterator.next();
			final String name = table.getTableName();
			final T reader = this.createNew(table);
			this.accessor.put(name, index, reader);
		}
	}

	public T getSpreadsheet(final int i) {
		final T spreadsheet;
		if (this.accessor.hasByIndex(i))
			spreadsheet = this.accessor.getByIndex(i);
		else {
			final List<Table> tables = this.document.getTableList();
			if (i < 0 || i >= tables.size())
				throw new NoSuchElementException(String.format(
						"No sheet at position %d", i));

			final Table table = this.document.getTableList().get(i);
			spreadsheet = this.createNew(table);
			this.accessor.put(table.getTableName(), i, spreadsheet);
		}
		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	protected T addSheetWithCheckedIndex(final int index, final String sheetName)
			throws CantInsertElementInSpreadsheetException {
		Table table = this.document.getTableByName(sheetName);
		if (table != null)
			throw new IllegalArgumentException(String.format("Sheet %s exists",
					sheetName));

		if (index == this.getSheetCount())
			table = Table.newTable(this.document);
		else
			table = this.document.insertSheet(index);
		if (table == null)
			throw new CantInsertElementInSpreadsheetException();

		final T spreadsheet = this.createNew(table);
		this.accessor.put(sheetName, index, spreadsheet);
		return spreadsheet;
	}

	protected abstract T createNew(
			/*>>> @UnknownInitialization AbstractOdsSimpleodfDocumentTrait<T> this, */final Table table);

	/** {@inheritDoc} */
	@Override
	protected T findSpreadsheetAndCreateReaderOrWriter(final String sheetName) {
		final T spreadsheet;
		final List<Table> tableList = this.document.getTableList();
		final ListIterator<Table> it = tableList.listIterator();
		Table table;
		while (it.hasNext()) {
			final int index = it.nextIndex();
			table = it.next();
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