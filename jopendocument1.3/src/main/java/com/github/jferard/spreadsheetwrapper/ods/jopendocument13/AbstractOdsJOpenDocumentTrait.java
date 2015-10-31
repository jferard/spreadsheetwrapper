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

import java.util.NoSuchElementException;

import org.jopendocument.dom.spreadsheet.Sheet;

import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentTrait;

/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

abstract class AbstractOdsJOpenDocumentTrait<T> extends
		AbstractSpreadsheetDocumentTrait<T> {
	private static int getSheetCount(
			final OdsJOpenStatefulDocument sfSpreadSheet) {
		int count;
		if (sfSpreadSheet.isNew())
			count = 0;
		else
			count = sfSpreadSheet.getRawSheetCount();
		return count;
	}

	/** the *internal* value (workbook) */
	private final OdsJOpenStatefulDocument sfSpreadSheet;

	/**
	 * @param spreadSheet
	 *            the *internal* value (workbook)
	 */
	public AbstractOdsJOpenDocumentTrait(
			final OdsJOpenStatefulDocument sfSpreadSheet) {
		super();
		this.sfSpreadSheet = sfSpreadSheet;
		final int sheetCount = AbstractOdsJOpenDocumentTrait
				.getSheetCount(sfSpreadSheet);
		for (int s = 0; s < sheetCount; s++) {
			final Sheet sheet = this.sfSpreadSheet.getRawSheet(s);
			final String name = sheet.getName();
			final T reader = this.createNew(sheet);
			this.accessor.put(name, s, reader);
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
			if (this.sfSpreadSheet.isNew() || index < 0
					|| index >= this.sfSpreadSheet.getRawSheetCount())
				throw new IndexOutOfBoundsException(String.format(
						"No sheet at position %d", index));

			final Sheet sheet = this.sfSpreadSheet.getRawSheet(index);
			spreadsheet = this.createNew(sheet);
			this.accessor.put(sheet.getName(), index, spreadsheet);
		}
		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	protected T addSheetWithCheckedIndex(final int index, final String sheetName) {
		Sheet sheet;
		if (this.sfSpreadSheet.isNew()
				&& this.sfSpreadSheet.getRawSheetCount() == 1) {
			sheet = this.sfSpreadSheet.getRawSheet(0);
			sheet.setName(sheetName);
		} else { // ok
			sheet = this.sfSpreadSheet.getRawSheet(sheetName);
			if (sheet != null)
				throw new IllegalArgumentException(String.format(
						"Sheet %s exists", sheetName));

			sheet = this.sfSpreadSheet.addRawSheet(index, sheetName);
		}
		this.sfSpreadSheet.setInitialized();
		final T spreadsheet = this.createNew(sheet);
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
			/*>>> @UnknownInitialization AbstractOdsJOpenDocumentTrait<T> this, */final Sheet sheet);

	/** {@inheritDoc} */
	@Override
	protected T findSpreadsheetAndCreateReaderOrWriter(final String sheetName) {
		if (this.sfSpreadSheet.isNew())
			throw new NoSuchElementException(String.format(
					"No %s sheet in workbook", sheetName));

		for (int s = 0; s < this.sfSpreadSheet.getRawSheetCount(); s++) {
			final Sheet sheet = this.sfSpreadSheet.getRawSheet(s);
			final String name = sheet.getName();
			if (name.equals(sheetName)) {
				final T spreadsheet = this.createNew(sheet);
				this.accessor.put(sheetName, s, spreadsheet);
				return spreadsheet;
			}
		}

		throw new NoSuchElementException(String.format(
				"No %s sheet in workbook", sheetName));
	}

	/** {@inheritDoc} */
	@Override
	protected final int getSheetCount() {
		return AbstractOdsJOpenDocumentTrait.getSheetCount(this.sfSpreadSheet);
	}
}