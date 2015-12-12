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

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.odftoolkit.odfdom.dom.style.OdfStyleFamily;
import org.odftoolkit.odfdom.incubator.doc.office.OdfOfficeStyles;
import org.odftoolkit.odfdom.incubator.doc.style.OdfStyle;
import org.odftoolkit.simple.table.Table;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.Output;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetWriterCursorImpl;
import com.github.jferard.spreadsheetwrapper.ods.odfdom.OdsOdfdomStyleHelper;

/*>>>
import org.checkerframework.checker.initialization.qual.UnknownInitialization;
import org.checkerframework.checker.nullness.qual.RequiresNonNull;
 */

/**
 *
 */
public class OdsSimpleodfDocumentWriter extends
AbstractSpreadsheetDocumentWriter implements SpreadsheetDocumentWriter {
	/** delegation value with definition of createNew */
	private final class OdsSimpleodfDocumentWriterTrait extends
	AbstractOdsSimpleodfDocumentTrait<SpreadsheetWriter> {

		/**
		 * @param traitStyleHelper
		 * @param value
		 *            *internal* workbook
		 */
		OdsSimpleodfDocumentWriterTrait(final OdsOdfdomStyleHelper styleHelper,
				final OdsSimpleodfStatefulDocument sfDocument) {
			super(sfDocument, styleHelper);
		}

		/** {@inheritDoc} */
		@Override
		/*@RequiresNonNull("traitStyleHelper")*/
		protected SpreadsheetWriter createNew(
				/*>>> @UnknownInitialization OdsSimpleodfDocumentWriterTrait this, */final Table table) {
			return new OdsSimpleodfWriter(table, this.traitStyleHelper);
		}
	}

	/** internal styles */
	private final OdfOfficeStyles documentStyles;

	/** for delegation */
	private final AbstractOdsSimpleodfDocumentTrait<SpreadsheetWriter> documentTrait;
	/** logger */
	private final Logger logger;

	/** reader for delegation */
	private final OdsSimpleodfDocumentReader reader;

	/** *internal* workbook */
	private final OdsSimpleodfStatefulDocument sfDocument;

	/** the style helper */
	private final OdsOdfdomStyleHelper styleHelper;

	/**
	 * @param logger
	 *            the logger
	 * @param value
	 *            *internal* workbook
	 * @param outputURL
	 *            URL where to write, should be a local file
	 * @throws SpreadsheetException
	 *             if the value writer can't be created
	 */
	public OdsSimpleodfDocumentWriter(final Logger logger,
			final OdsOdfdomStyleHelper styleHelper,
			final OdsSimpleodfStatefulDocument sfDocument, final Output output)
					throws SpreadsheetException {
		super(logger, output);
		this.styleHelper = styleHelper;
		this.reader = new OdsSimpleodfDocumentReader(styleHelper, sfDocument);
		this.logger = logger;
		this.sfDocument = sfDocument;
		this.documentStyles = this.sfDocument.getStyles();
		this.documentTrait = new OdsSimpleodfDocumentWriterTrait(styleHelper,
				sfDocument);
	}

	/**
	 * {@inheritDoc}
	 *
	 * @throws CantInsertElementInSpreadsheetException
	 * @throws IndexOutOfBoundsException
	 */
	@Override
	public SpreadsheetWriter addSheet(final int index, final String sheetName)
			throws IndexOutOfBoundsException,
			CantInsertElementInSpreadsheetException {
		return this.documentTrait.addSheet(index, sheetName);
	}

	/**
	 * {@inheritDoc}
	 *
	 * @throws CantInsertElementInSpreadsheetException
	 */
	@Override
	public SpreadsheetWriter addSheet(final String sheetName)
			throws CantInsertElementInSpreadsheetException {
		return this.documentTrait.addSheet(sheetName);
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		try {
			this.output.close();
		} catch (final IOException e) {
			final String message = e.getMessage();
			this.logger.log(Level.SEVERE, message == null ? "" : message, e);
		}
		this.reader.close();
	}

	/** {@inheritDoc} */
	@Override
	public WrapperCellStyle getCellStyle(final String styleName) {
		return this.reader.getCellStyle(styleName);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriterCursor getNewCursorByIndex(final int index) {
		return new SpreadsheetWriterCursorImpl(this.getSpreadsheet(index));
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriterCursor getNewCursorByName(final String sheetName)
			throws SpreadsheetException {
		return new SpreadsheetWriterCursorImpl(this.getSpreadsheet(sheetName));
	}

	/** {@inheritDoc} */
	@Override
	public int getSheetCount() {
		return this.reader.getSheetCount();
	}

	/** {@inheritDoc} */
	@Override
	public List<String> getSheetNames() {
		return this.reader.getSheetNames();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter getSpreadsheet(final int index) {
		return this.documentTrait.getSpreadsheet(index);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter getSpreadsheet(final String sheetName) {
		return this.documentTrait.getSpreadsheet(sheetName);
	}

	/** {@inheritDoc} */
	@Override
	public void save() throws SpreadsheetException {
		OutputStream outputStream = null;
		try {
			outputStream = this.output.getStream();
			if (outputStream == null)
				throw new IllegalStateException(
						String.format("Use saveAs when output file is not specified"));
			this.sfDocument.rawSave(outputStream);
		} catch (final Exception e) { // NOPMD by Julien on 02/09/15 22:55
			this.logger.log(Level.SEVERE, String.format(
					"this.spreadsheetDocument.save(%s) not ok", outputStream),
					e);
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyle(final String styleName,
			final WrapperCellStyle wrapperCellStyle) {
		final OdfStyle newStyle = this.documentStyles.newStyle(styleName,
				OdfStyleFamily.TableCell);
		newStyle.setProperties(this.styleHelper.getProperties(wrapperCellStyle));
		return true;
	}
}
