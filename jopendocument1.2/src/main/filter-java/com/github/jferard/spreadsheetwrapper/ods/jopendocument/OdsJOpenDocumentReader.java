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

package com.github.jferard.spreadsheetwrapper.ods.${jopendocument.pkg};

import java.util.ArrayList;
import java.util.List;

import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReaderCursor;
import com.github.jferard.spreadsheetwrapper.Stateful;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetReaderCursorImpl;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

/**
 */
class OdsJOpenDocumentReader implements SpreadsheetDocumentReader {
	/** delegation value with definition of createNew */
	private final class OdsJOpenDocumentReaderDelegate extends
	AbstractOdsJOpenDocumentDelegate<SpreadsheetReader> {
		/** the style helper */
		private final OdsJOpenStyleHelper styleHelper;

		/**
		 * @param styleHelper
		 * 				the style helper
		 * @param sfSpreadSheet
		 *            *internal* value
		 */
		OdsJOpenDocumentReaderDelegate(final OdsJOpenStyleHelper styleHelper,
				final OdsJOpenStatefulDocument sfSpreadSheet) {
			super(sfSpreadSheet);
			this.styleHelper = styleHelper;
		}

		/** {@inheritDoc} */
		@Override
		protected SpreadsheetReader createNew(
				/*>>> @UnknownInitialization OdsJOpenDocumentReaderDelegate this,*/final Sheet sheet) {
			return new OdsJOpenReader(this.styleHelper, sheet);
		}
	}

	/** for delegation only */
	private final AbstractOdsJOpenDocumentDelegate<SpreadsheetReader> documentDelegate;
	/** the *internal* value */
	private final Stateful<SpreadSheet> sfSpreadSheet;

	/**
	 * @param styleHelper
	 * 				the style helper
	 * @param sfSpreadSheet
	 *            *internal* value
	 * @throws SpreadsheetException
	 *             if can't open reader
	 */
	OdsJOpenDocumentReader(final OdsJOpenStyleHelper styleHelper,
			final OdsJOpenStatefulDocument sfSpreadSheet)
					throws SpreadsheetException {
		this.sfSpreadSheet = sfSpreadSheet;
		this.documentDelegate = new OdsJOpenDocumentReaderDelegate(styleHelper,
				sfSpreadSheet);
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		// do nothing ??
	}

	/** {@inheritDoc} */
	@Override
	public WrapperCellStyle getCellStyle(final String styleName) {
		throw new UnsupportedOperationException();
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
		return this.documentDelegate.getSheetCount();
	}

	/** */
	@Override
	public List<String> getSheetNames() {
		final List<String> sheetNames;
		final int sheetCount = this.getSheetCount();
		sheetNames = new ArrayList<String>(sheetCount);
		for (int s = 0; s < sheetCount; s++)
			sheetNames
			.add(this.sfSpreadSheet.getObject().getSheet(s).getName());
		return sheetNames;
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetReader getSpreadsheet(final int index) {
		return this.documentDelegate.getSpreadsheet(index);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetReader getSpreadsheet(final String sheetName) {
		return this.documentDelegate.getSpreadsheet(sheetName);
	}
}