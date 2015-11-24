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
package com.github.jferard.spreadsheetwrapper.ods.simpleods;

import java.util.NoSuchElementException;

import org.simpleods.OdsFile;
import org.simpleods.SimpleOdsException;
import org.simpleods.Table;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentTrait;

/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

/**
 * @param <T>
 *            reader/writer
 */
public abstract class AbstractOdsSimpleodsDocumentTrait<T> extends
AbstractSpreadsheetDocumentTrait<T> {
	/** the *internal* ods file */
	private final OdsFile file;

	/**
	 * Creates a value from an ods file
	 *
	 * @param file
	 *            the *internal* ods file
	 */
	public AbstractOdsSimpleodsDocumentTrait(final OdsFile file) {
		super();
		this.file = file;
		final int count = this.getSheetCount();
		try {
			for (int i = 0; i < count; i++) {
				// SimpleOdsException
				// if i < 0 or i
				// >= table
				// count
				final Table table = this.getTable(i);
				if (table == null)
					throw new AssertionError();

				final T reader = this.createNew(table);
				final String name = this.file.getTableName(i); // throws
				this.accessor.put(name, i, reader);
			}
		} catch (final SimpleOdsException e) {
			throw new AssertionError(e);
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
			final int count = this.getSheetCount();
			if (index < 0 || index >= count)
				throw new NoSuchElementException(String.format(
						"No sheet at position %d", index));

			try {
				final String name = this.file.getTableName(index);
				final Table table = this.getTable(index);
				spreadsheet = this.createNew(table);
				this.accessor.put(name, index, spreadsheet);
			} catch (final SimpleOdsException e) {
				throw new AssertionError(String.format(
						"Can't reach sheet at position %d", index), e);
			}
		}
		return spreadsheet;
	}

	private Table getTable(
			/*>>> @UnknownInitialization AbstractOdsSimpleodsDocumentTrait<T> this, */final int index) {
		if (this.file == null)
			throw new IllegalStateException();

		return (Table) this.file.getContent().getTableQueue().get(index);
	}

	/** {@inheritDoc} */
	@Override
	protected T addSheetWithCheckedIndex(final int index, final String sheetName)
			throws CantInsertElementInSpreadsheetException {
		if (index != this.getSheetCount())
			throw new UnsupportedOperationException();

		final int indexForSheetName = this.file.getTableNumber(sheetName);
		if (indexForSheetName != -1)
			throw new IllegalArgumentException(String.format("Sheet %s exists",
					sheetName));

		// at the end
		final T spreadsheet;
		try {
			this.file.addTable(sheetName);
			final Table table = this.getTable(index);
			spreadsheet = this.createNew(table);
			this.accessor.put(sheetName, index, spreadsheet);
		} catch (final SimpleOdsException e) {
			throw new CantInsertElementInSpreadsheetException(e);
		}
		return spreadsheet;
	}

	/**
	 * @param table
	 *            *internal* sheet
	 * @return the reader/writer
	 */
	protected abstract T createNew(
			/*>>> @UnknownInitialization AbstractOdsSimpleodsDocumentTrait<T> this, */final Table table);

	/** {@inheritDoc} */
	@Override
	protected T findSpreadsheetAndCreateReaderOrWriter(final String sheetName) {
		final int index = this.file.getTableNumber(sheetName);
		if (index == -1)
			throw new NoSuchElementException(String.format(
					"No %s sheet in workbook", sheetName));

		final Table table = this.getTable(index);
		final T spreadsheet = this.createNew(table);
		this.accessor.put(sheetName, index, spreadsheet);
		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	protected final int getSheetCount(/*>>> @UnknownInitialization(AbstractOdsSimpleodsDocumentTrait.class) AbstractOdsSimpleodsDocumentTrait<T> this*/) {
		if (this.file == null)
			throw new IllegalStateException();
		return this.file.lastTableNumber();
	}
}