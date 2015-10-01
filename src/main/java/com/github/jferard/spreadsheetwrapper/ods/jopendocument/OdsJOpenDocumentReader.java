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

package com.github.jferard.spreadsheetwrapper.ods.jopendocument;

import java.util.ArrayList;
import java.util.List;

import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReaderCursor;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetReaderCursorImpl;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/
/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

/**
 */
class OdsJOpenDocumentReader implements SpreadsheetDocumentReader {
	/** delegation document with definition of createNew */
	private final class OdsJOpenDocumentReaderTrait extends
			AbstractOdsJOpenDocumentTrait<SpreadsheetReader> {
		/**
		 * @param spreadSheet
		 *            *internal* document
		 */
		OdsJOpenDocumentReaderTrait(final SpreadSheet spreadSheet) {
			super(spreadSheet);
		}

		/** {@inheritDoc} */
		@Override
		protected SpreadsheetReader createNew(
				/*>>> @UnknownInitialization OdsJOpenDocumentReaderTrait this,*/final Sheet sheet) {
			return new OdsJOpenReader(sheet);
		}
	}

	/** the *internal* document */
	private final SpreadSheet spreadSheet;
	/** for delegation only */
	private final AbstractOdsJOpenDocumentTrait<SpreadsheetReader> spreadSheetTrait;

	/**
	 * @param spreadSheet
	 *            *internal* document
	 * @throws SpreadsheetException
	 *             if can't open reader
	 */
	OdsJOpenDocumentReader(final SpreadSheet spreadSheet)
			throws SpreadsheetException {
		this.spreadSheet = spreadSheet;
		this.spreadSheetTrait = new OdsJOpenDocumentReaderTrait(spreadSheet);
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		// do nothing ??
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

	/** */
	@Override
	public int getSheetCount() {
		return this.spreadSheet.getSheetCount();
	}

	/** */
	@Override
	public List<String> getSheetNames() {
		int sheetCount = this.getSheetCount();
		final List<String> tableNames = new ArrayList<String>(sheetCount);
		for (int s = 0; s<sheetCount; s++)
			tableNames.add(this.spreadSheet.getSheet(s).getName());
		return tableNames;
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetReader getSpreadsheet(final int index) {
		return this.spreadSheetTrait.getSpreadsheet(index);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetReader getSpreadsheet(final String sheetName) {
		return this.spreadSheetTrait.getSpreadsheet(sheetName);
	}
	
	
}