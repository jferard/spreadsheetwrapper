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

package com.github.jferard.spreadsheetwrapper.ods.odfdom;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;
import org.odftoolkit.odfdom.doc.table.OdfTable;
import org.odftoolkit.odfdom.dom.style.OdfStyleFamily;
import org.odftoolkit.odfdom.dom.style.props.OdfTableCellProperties;
import org.odftoolkit.odfdom.dom.style.props.OdfTextProperties;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;

import com.github.jferard.spreadsheetwrapper.CellStyle;
import com.github.jferard.spreadsheetwrapper.CellStyle.Color;
import com.github.jferard.spreadsheetwrapper.Font;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReaderCursor;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetReaderCursorImpl;
import com.github.jferard.spreadsheetwrapper.ods.OdsUtility;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/
/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

/**
 */
class OdsOdfdomDocumentReader implements SpreadsheetDocumentReader {
	/** delegation value with definition of createNew */
	private final class OdsOdfdomDocumentReaderTrait extends
			AbstractOdsOdfdomDocumentTrait<SpreadsheetReader> {
		/**
		 * @param value
		 *            *internal* value
		 */
		OdsOdfdomDocumentReaderTrait(final OdfSpreadsheetDocument document) {
			super(document);
		}

		/** {@inheritDoc} */
		@Override
		protected SpreadsheetReader createNew(
				/*>>> @UnknownInitialization OdsOdfdomDocumentReaderTrait this,*/final OdfTable table) {
			return new OdsOdfdomReader(table);
		}
	}

	/** the *internal* value */
	private final OdfSpreadsheetDocument document;
	/** for delegation only */
	private final AbstractOdsOdfdomDocumentTrait<SpreadsheetReader> documentTrait;
	private OdfOfficeStyles documentStyles;

	/**
	 * @param value
	 *            *internal* value
	 * @throws SpreadsheetException
	 *             if can't open reader
	 */
	OdsOdfdomDocumentReader(final OdfSpreadsheetDocument document)
			throws SpreadsheetException {
		this.document = document;
		try {
			this.document.getContentRoot();
		} catch (final Exception e) { // NOPMD by Julien on 03/09/15 18:39
			throw new SpreadsheetException(e); // colors the exception
		}
		this.document.setLocale(Locale.FRANCE);
		this.documentStyles = this.document.getDocumentStyles();
		this.documentTrait = new OdsOdfdomDocumentReaderTrait(document);
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		this.document.close();
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

	// private SpreadsheetReaderCursor getPrivateCursorByIndex(final int i)
	// throws SpreadsheetException {
	// final SpreadsheetWriterCursor cursor;
	// final List<OdfTable> sheets = this.document.getTableList();
	// if (i < 0 || i >= sheets.size())
	// throw new SpreadsheetException(
	// String.format("No sheet at index %d", i));
	// else
	// cursor = this.getNewCursorByIndex(i);
	// return cursor;
	// }

	/** */
	@Override
	public int getSheetCount() {
		final List<OdfTable> tables = this.document.getTableList();
		return tables.size();
	}

	// private SpreadsheetReaderCursor getPrivateCursorByName(final String
	// sheetName)
	// throws SpreadsheetException {
	// final SpreadsheetWriterCursor cursor;
	// if (this.cursorHandler.has(sheetName))
	// cursor = this.cursorHandler.get(sheetName);
	// else
	// cursor = this.getNewCursorByName(sheetName);
	// return cursor;
	// }

	/** */
	@Override
	public List<String> getSheetNames() {
		final List<OdfTable> tables = this.document.getTableList();
		final List<String> tableNames = new ArrayList<String>(tables.size());
		for (final OdfTable table : tables)
			tableNames.add(table.getTableName());
		return tableNames;
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
	
	/** {@inheritDoc} */
	@Override
	public CellStyle getCellStyle(String styleName) {
		OdfStyle existingStyle = this.documentStyles.getStyle(styleName, OdfStyleFamily.TableCell);
		return OdsUtility.getCellStyle(existingStyle);
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public String getStyleString(String styleName) {
		OdfStyle existingStyle = this.documentStyles.getStyle(styleName, OdfStyleFamily.TableCell);
		return OdsUtility.getStyleString(existingStyle);
	}
}