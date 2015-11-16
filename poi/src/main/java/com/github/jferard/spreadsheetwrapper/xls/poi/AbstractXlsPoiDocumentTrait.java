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
package com.github.jferard.spreadsheetwrapper.xls.poi;

import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentTrait;

/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/
/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/
/*>>> import org.checkerframework.checker.nullness.qual.PolyNull;*/

public abstract class AbstractXlsPoiDocumentTrait<T> extends
		AbstractSpreadsheetDocumentTrait<T> {
	/** *internal* workbook */
	final private Workbook workbook;
	/** cell style for date cells, since Excel hasn't any cell date type */
	protected/*@Nullable*/CellStyle dateCellStyle;
	protected XlsPoiStyleHelper styleHelper;

	/**
	 * @param workbook
	 *            *internal* workbook
	 * @param styleHelper 
	 * @param dateCellStyle
	 *            cell syle for dates, since Excel hasn't a real date type
	 */
	public AbstractXlsPoiDocumentTrait(final Workbook workbook, /*@Nullable*/
			XlsPoiStyleHelper styleHelper, final CellStyle dateCellStyle) {
		super();
		this.workbook = workbook;
		this.styleHelper = styleHelper;
		this.dateCellStyle = dateCellStyle;
		final int sheetCount = this.getSheetCount();
		for (int i = 0; i < sheetCount; i++) {
			final Sheet sheet = this.workbook.getSheetAt(i);
			final String name = sheet.getSheetName();
			final T reader = this.createNew(sheet);
			this.accessor.put(name, i, reader);
		}
	}

	/**
	 * @param index
	 *            index of the sheet in the workbook
	 * @return the reader/writer on the sheet
	 */
	public T getSpreadsheet(final int index) {
		final T spreadsheet;
		if (this.accessor.hasByIndex(index))
			spreadsheet = this.accessor.getByIndex(index);
		else {
			if (index < 0 || index >= this.getSheetCount())
				throw new IndexOutOfBoundsException(String.format(
						"No sheet at position %d", index));

			final Sheet sheet = this.workbook.getSheetAt(index);
			spreadsheet = this.createNew(sheet);
			this.accessor.put(sheet.getSheetName(), index, spreadsheet);
		}
		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	protected T addSheetWithCheckedIndex(final int index, final String sheetName) {
		Sheet sheet = this.workbook.getSheet(sheetName);
		if (sheet != null)
			throw new IllegalArgumentException(String.format("Sheet %s exists",
					sheetName));

		sheet = this.workbook.createSheet(sheetName);
		this.workbook.setSheetOrder(sheetName, index);
		final T spreadsheet = this.createNew(sheet);
		this.accessor.put(sheetName, index, spreadsheet);
		return spreadsheet;
	}

	/**
	 * @param sheet
	 *            the *internal* sheet
	 * @return reader/writer on that sheet
	 */
	protected abstract T createNew(
			/*>>> @UnknownInitialization AbstractXlsPoiDocumentTrait<T> this, */final Sheet sheet);

	/** {@inheritDoc} */
	@Override
	protected T findSpreadsheetAndCreateReaderOrWriter(final String sheetName) {
		final Sheet sheet = this.workbook.getSheet(sheetName);
		if (sheet == null)
			throw new NoSuchElementException(String.format(
					"No %s sheet in workbook", sheetName));

		final int index = this.workbook.getSheetIndex(sheet);
		final T spreadsheet = this.createNew(sheet);
		this.accessor.put(sheetName, index, spreadsheet);
		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	protected final int getSheetCount(/*>>> @UnknownInitialization AbstractXlsPoiDocumentTrait<T> this*/) {
		if (this.workbook == null)
			throw new IllegalStateException();
		return this.workbook.getNumberOfSheets();
	}
}