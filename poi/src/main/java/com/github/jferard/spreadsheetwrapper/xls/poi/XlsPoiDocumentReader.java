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

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReaderCursor;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetReaderCursorImpl;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>>
import org.checkerframework.checker.nullness.qual.Nullable;
import org.checkerframework.checker.initialization.qual.UnknownInitialization;
import org.checkerframework.checker.nullness.qual.RequiresNonNull;
 */

/**
 */
public class XlsPoiDocumentReader implements SpreadsheetDocumentReader {
	/** for delegation */
	private static final class XlsPoiDocumentReaderTrait extends
			AbstractXlsPoiDocumentTrait<SpreadsheetReader> {

		XlsPoiDocumentReaderTrait(final Workbook workbook,
				final XlsPoiStyleHelper styleHelper) {
			super(workbook, styleHelper, null);
		}

		/** {inheritDoc} */
		@Override
		/*@RequiresNonNull("traitStyleHelper")*/
		protected SpreadsheetReader createNew(
				/*>>> @UnknownInitialization XlsPoiDocumentReaderTrait this, */final Sheet sheet) {
			return new XlsPoiReader(sheet, this.traitStyleHelper);
		}
	}

	/** for delegation */
	private final AbstractXlsPoiDocumentTrait<SpreadsheetReader> documentTrait;

	/** for delegation */
	private final XlsPoiStyleHelper styleHelper;
	/** *internal* workbook */
	private final Workbook workbook;

	/**
	 * @param workbook
	 *            *internal* workbook
	 * @param cellStyleAccessor
	 */
	XlsPoiDocumentReader(final Workbook workbook,
			final XlsPoiStyleHelper styleHelper) {
		this.workbook = workbook;
		this.styleHelper = styleHelper;
		this.documentTrait = new XlsPoiDocumentReaderTrait(workbook,
				styleHelper);
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		// Nothing to do
	}

	/** {@inheritDoc} */
	@Override
	public/*@Nullable*/WrapperCellStyle getCellStyle(final String styleName) {
		final CellStyle cellStyle = this.styleHelper.getCellStyle(
				this.workbook, styleName);
		if (cellStyle == null)
			return null;

		return this.styleHelper.toWrapperCellStyle(this.workbook, cellStyle);
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

	/** {@inheritDoc} */
	@Override
	public int getSheetCount() {
		return this.workbook.getNumberOfSheets();
	}

	/** {@inheritDoc} */
	@Override
	public List<String> getSheetNames() {
		final int sheetCount = this.getSheetCount();
		final List<String> sheetNames = new ArrayList<String>(sheetCount);
		for (int i = 0; i < sheetCount; i++)
			sheetNames.add(this.workbook.getSheetName(i));
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