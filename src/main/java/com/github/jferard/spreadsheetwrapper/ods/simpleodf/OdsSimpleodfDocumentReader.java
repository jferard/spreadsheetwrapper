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

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.table.Table;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReaderCursor;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetReaderCursorImpl;

/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

/**
 * The document reader for Apache simple-odf
 */
public class OdsSimpleodfDocumentReader implements SpreadsheetDocumentReader {
	/** delegation document with definition of createNew */
	private final class OdsSimpleodfDocumentWriterTrait extends
			AbstractOdsSimpleodfDocumentTrait<SpreadsheetReader> {
		OdsSimpleodfDocumentWriterTrait(final SpreadsheetDocument document) {
			super(document);
		}

		/** {@inheritDoc} */
		@Override
		protected SpreadsheetReader createNew(
				/*>>> @UnknownInitialization OdsSimpleodfDocumentWriterTrait this, */final Table table) {
			return new OdsSimpleodfReader(table);
		}
	}

	/** *internal* workbook */
	private final SpreadsheetDocument document;
	/** for delegation */
	private final AbstractOdsSimpleodfDocumentTrait<SpreadsheetReader> documentTrait;

	/**
	 * @param document
	 *            *internal* workbook
	 * @throws SpreadsheetException
	 *             if the document reader can't be created
	 */
	OdsSimpleodfDocumentReader(final SpreadsheetDocument document)
			throws SpreadsheetException {
		this.document = document;
		try {
			this.document.getContentRoot();
		} catch (final Exception e) { // NOPMD by Julien on 03/09/15 22:09
			throw new SpreadsheetException(e); // colors the exception
		}
		this.document.setLocale(Locale.US);
		this.documentTrait = new OdsSimpleodfDocumentWriterTrait(document);
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		this.document.close();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetReaderCursor getNewCursorByIndex(final int index)
			throws SpreadsheetException {
		return new SpreadsheetReaderCursorImpl(this.getSpreadsheet(index));
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetReaderCursor getNewCursorByName(final String sheetName)
			throws SpreadsheetException {
		return new SpreadsheetReaderCursorImpl(this.getSpreadsheet(sheetName));
	}

	/** {@inheritDoc} */
	@Override
	public int getSheetCount() {
		final List<Table> tables = this.document.getTableList();
		return tables.size();
	}

	/** {@inheritDoc} */
	@Override
	public List<String> getSheetNames() {
		final List<Table> tables = this.document.getTableList();
		final List<String> sheetNames = new ArrayList<String>(tables.size());
		for (final Table table : tables)
			sheetNames.add(table.getTableName());
		return sheetNames;
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetReader getSpreadsheet(final int index) {
		return this.documentTrait.getSpreadsheet(index);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetReader getSpreadsheet(final String sheetName) {
		return this.documentTrait.getSpreadsheet(sheetName);
	}
}