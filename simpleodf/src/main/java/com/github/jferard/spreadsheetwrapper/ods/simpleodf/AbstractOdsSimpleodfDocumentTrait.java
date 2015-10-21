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

import java.util.Collections;
import java.util.List;
import java.util.ListIterator;
import java.util.NoSuchElementException;

import org.odftoolkit.odfdom.dom.element.table.TableTableColumnElement;
import org.odftoolkit.odfdom.dom.element.table.TableTableElement;
import org.odftoolkit.simple.table.Table;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentTrait;

/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

public abstract class AbstractOdsSimpleodfDocumentTrait<T> extends
		AbstractSpreadsheetDocumentTrait<T> {
	/** the value wrapper for delegation */
	private final OdsSimpleodfStatefulDocument sfDocument;

	/**
	 * @param sfDocument the stateful (ie inited/not inited) document
	 */
	public AbstractOdsSimpleodfDocumentTrait(
			final OdsSimpleodfStatefulDocument sfDocument) {
		super();
		this.sfDocument = sfDocument;
		final List<Table> tables = this.getTableList();
		final ListIterator<Table> tablesIterator = tables.listIterator();
		while (tablesIterator.hasNext()) {
			final Integer index = tablesIterator.nextIndex();
			final Table table = tablesIterator.next();
			final String name = table.getTableName();
			final T reader = this.createNew(table);
			this.accessor.put(name, index, reader);
		}
	}

	public T getSpreadsheet(final int index) {
		final T spreadsheet;
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

	public final List<Table> getTableList() {
		final List<Table> tables;
		if (this.sfDocument.isNew())
			tables = Collections.emptyList();
		else
			tables = this.sfDocument.getRawTableList();
		return tables;
	}

	/** {@inheritDoc} */
	@Override
	protected T addSheetWithCheckedIndex(final int index, final String sheetName)
			throws CantInsertElementInSpreadsheetException {
		Table table = this.sfDocument.getRawSheet(sheetName);
		if (table != null)
			throw new IllegalArgumentException(String.format("Sheet %s exists",
					sheetName));

		if (this.sfDocument.isNew()
				&& this.sfDocument.getRawTableList().size() >= 1) {
			table = this.sfDocument.getRawSheet(0);
			table.setTableName(sheetName);
		} else {
			if (index == this.getSheetCount())
				table = this.sfDocument.rawNewTable();
			else
				table = this.sfDocument.rawInsertSheet(index);
			if (table == null)
				throw new CantInsertElementInSpreadsheetException();
		}
		TableTableElement tableElement = table.getOdfElement();
		this.cleanEmptyTable(tableElement);

		this.sfDocument.setInitialized();
		final T spreadsheet = this.createNew(table);
		this.accessor.put(sheetName, index, spreadsheet);
		return spreadsheet;
	}
	
	private void cleanEmptyTable(final TableTableElement tableElement) {
		final NodeList colsList = tableElement
				.getElementsByTagName("table:table-column");
		assert colsList.getLength() == 1;
		TableTableColumnElement column = (TableTableColumnElement) colsList.item(0);
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
	

	protected abstract T createNew(
			/*>>> @UnknownInitialization AbstractOdsSimpleodfDocumentTrait<T> this, */final Table table);

	/** {@inheritDoc} */
	@Override
	protected T findSpreadsheetAndCreateReaderOrWriter(final String sheetName) {
		final T spreadsheet;
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

	/** {@inheritDoc} */
	@Override
	protected int getSheetCount() {

		return this.getTableList().size();
	}
}