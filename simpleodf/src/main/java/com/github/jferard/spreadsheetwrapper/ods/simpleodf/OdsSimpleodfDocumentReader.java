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

import org.odftoolkit.odfdom.dom.style.OdfStyleFamily;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;
import org.odftoolkit.simple.table.Table;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReaderCursor;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetReaderCursorImpl;
import com.github.jferard.spreadsheetwrapper.ods.odfdom.OdsOdfdomStyleHelper;

/*>>>
import org.checkerframework.checker.initialization.qual.UnknownInitialization;
import org.checkerframework.checker.nullness.qual.RequiresNonNull;
 */

/**
 * The value reader for Apache simple-odf
 */
public class OdsSimpleodfDocumentReader implements SpreadsheetDocumentReader {
	/** delegation value with definition of createNew */
	private final class OdsSimpleodfDocumentReaderTrait extends
			AbstractOdsSimpleodfDocumentTrait<SpreadsheetReader> {

		OdsSimpleodfDocumentReaderTrait(
				final OdsSimpleodfStatefulDocument sfDocument,
				final OdsOdfdomStyleHelper styleHelper) {
			super(sfDocument, styleHelper);
		}

		/** {@inheritDoc} */
		@Override
		/*@RequiresNonNull("traitStyleHelper")*/
		protected SpreadsheetReader createNew(
				/*>>> @UnknownInitialization OdsSimpleodfDocumentReaderTrait this, */final Table table) {
			return new OdsSimpleodfReader(table, this.traitStyleHelper);
		}
	}

	/** internal styles */
	private final OdfOfficeStyles documentStyles;
	/** for delegation */
	private final AbstractOdsSimpleodfDocumentTrait<SpreadsheetReader> documentTrait;
	/** *internal* workbook */
	private final OdsSimpleodfStatefulDocument sfDocument;
	/** the style helper */
	private final OdsOdfdomStyleHelper styleHelper;

	/**
	 * @param value
	 *            *internal* workbook
	 * @throws SpreadsheetException
	 *             if the value reader can't be created
	 */
	OdsSimpleodfDocumentReader(final OdsOdfdomStyleHelper styleHelper,
			final OdsSimpleodfStatefulDocument sfDocument)
					throws SpreadsheetException {
		this.styleHelper = styleHelper;
		this.sfDocument = sfDocument;
		// try {
		// this.sfDocument.getValue().getContentRoot();
		// } catch (final Exception e) { // NOPMD by Julien on 03/09/15 22:09
		// throw new SpreadsheetException(e); // colors the exception
		// }
		this.sfDocument.rawSetLocale(Locale.US);
		this.documentStyles = this.sfDocument.getStyles();
		this.documentTrait = new OdsSimpleodfDocumentReaderTrait(sfDocument,
				styleHelper);
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		this.sfDocument.close();
	}

	/** {@inheritDoc} */
	@Override
	public WrapperCellStyle getCellStyle(final String styleName) {
		final OdfStyle existingStyle = this.documentStyles.getStyle(styleName,
				OdfStyleFamily.TableCell);
		return this.styleHelper.toCellStyle(existingStyle);
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
		final List<Table> tables = this.documentTrait.getTableList();
		return tables.size();
	}

	/** {@inheritDoc} */
	@Override
	public List<String> getSheetNames() {
		final List<Table> tables = this.documentTrait.getTableList();
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