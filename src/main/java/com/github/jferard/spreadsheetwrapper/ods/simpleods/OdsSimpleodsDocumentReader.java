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

import java.util.ArrayList;
import java.util.List;

import org.simpleods.OdsFile;
import org.simpleods.SimpleOdsException;
import org.simpleods.Table;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReaderCursor;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetReaderCursorImpl;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/
/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

/**
 */
public class OdsSimpleodsDocumentReader implements SpreadsheetDocumentReader {
	/** a document for delegation */
	private final class OdsSimpleodsDocumentReaderTrait extends
			AbstractOdsSimpleodsDocumentTrait<SpreadsheetReader> {
		/**
		 * @param file
		 *            *internal* workbook
		 */
		OdsSimpleodsDocumentReaderTrait(final OdsFile file) {
			super(file);
		}

		/** {@inheritDoc} */
		@Override
		protected SpreadsheetReader createNew(
				/*>>> @UnknownInitialization OdsSimpleodsDocumentReaderTrait this, */final Table table) {
			return new OdsSimpleodsReader(table);
		}
	}

	/** the document for delegation */
	private final AbstractOdsSimpleodsDocumentTrait<SpreadsheetReader> documentTrait;
	/** *internal* workbook */
	private final OdsFile file;

	/**
	 * @param file
	 *            *internal* workbook
	 */
	OdsSimpleodsDocumentReader(final OdsFile file) {
		this.file = file;
		this.documentTrait = new OdsSimpleodsDocumentReaderTrait(file);
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		// this.file.close();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetReaderCursor getNewCursorByIndex(final int index) {
		return new SpreadsheetReaderCursorImpl(this.getSpreadsheet(index));
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetReaderCursor getNewCursorByName(final String sheetName)
			throws SpreadsheetException {
		return new SpreadsheetReaderCursorImpl(this.getSpreadsheet(sheetName));
	}

	/** */
	@Override
	public int getSheetCount() {
		return this.file.lastTableNumber();
	}

	/** */
	@Override
	public List<String> getSheetNames() {
		final int count = this.getSheetCount();
		final List<String> sheetNames = new ArrayList<String>(count);
		try {
			for (int i = 0; i < count; i++)
				sheetNames.add(this.file.getTableName(i));
		} catch (final SimpleOdsException e) {
			throw new AssertionError(e);
		}
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