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

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.jdom.Element;
import org.jopendocument.dom.ODXMLDocument;
import org.jopendocument.dom.spreadsheet.Sheet;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.Output;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetWriterCursorImpl;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/
/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * The SpreadSheetWriter for JOpendocument
 *
 */
class OdsJOpenDocumentWriter extends AbstractSpreadsheetDocumentWriter
implements SpreadsheetDocumentWriter {
	/** delegation sfSpreadSheet with definition of createNew */
	private final class OdsJOpenDocumentWriterDelegate extends
	AbstractOdsJOpenDocumentDelegate<SpreadsheetWriter> {
		/** the style helper */
		private final OdsJOpenStyleHelper styleHelper;

		/**
		 * @param styleHelper
		 * 			the style helper
		 * @param sfSpreadSheet
		 * 			the *internal* spreadsheet, stateful version
		 */
		OdsJOpenDocumentWriterDelegate(final OdsJOpenStyleHelper styleHelper,
				final OdsJOpenStatefulDocument sfSpreadSheet) {
			super(sfSpreadSheet);
			this.styleHelper = styleHelper;
		}

		/** {@inheritDoc} */
		@Override
		protected SpreadsheetWriter createNew(
				/*>>> @UnknownInitialization OdsJOpenDocumentWriterDelegate this, */final Sheet sheet) {
			return new OdsJOpenWriter(this.styleHelper, sheet);
		}
	}

	/** delegation sfSpreadSheet */
	private final AbstractOdsJOpenDocumentDelegate<SpreadsheetWriter> documentDelegate;
	/** the logger */
	private final Logger logger;
	/** delegation reader */
	private final OdsJOpenDocumentReader reader;

	/** the *internal* workbook */
	private final OdsJOpenStatefulDocument sfSpreadSheet;

	/** utility for style */
	private final OdsJOpenStyleHelper styleHelper;

	/**
	 * @param logger
	 *            the logger
	 * @param styleHelper
	 *            the basic style helper
	 * @param sfSpreadSheet
	 *            the *internal* workbook
	 * @param output
	 *            where to write
	 * @param newDocument
	 * @throws SpreadsheetException
	 *             if can't open sfSpreadSheet writer
	 */
	OdsJOpenDocumentWriter(final Logger logger,
			final OdsJOpenStyleHelper styleHelper,
			final OdsJOpenStatefulDocument sfSpreadSheet, final Output output)
					throws SpreadsheetException {
		super(logger, output);
		this.styleHelper = styleHelper;
		this.reader = new OdsJOpenDocumentReader(styleHelper, sfSpreadSheet);
		this.logger = logger;
		this.sfSpreadSheet = sfSpreadSheet;
		this.documentDelegate = new OdsJOpenDocumentWriterDelegate(styleHelper, sfSpreadSheet);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter addSheet(final int index, final String sheetName)
			throws IndexOutOfBoundsException,
			CantInsertElementInSpreadsheetException {
		return this.documentDelegate.addSheet(index, sheetName);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter addSheet(final String sheetName)
			throws CantInsertElementInSpreadsheetException {
		return this.documentDelegate.addSheet(sheetName);
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
		return this.documentDelegate.getSpreadsheet(index);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter getSpreadsheet(final String sheetName) {
		return this.documentDelegate.getSpreadsheet(sheetName);
	}

	/** */
	@Override
	public void save() throws SpreadsheetException {
		OutputStream outputStream = null;
		try {
			outputStream = this.output.getStream();
			if (outputStream == null)
				throw new IllegalStateException(
						String.format("Use saveAs when output file/stream is not specified at opening"));

			this.sfSpreadSheet.save(outputStream);
		} catch (final Exception e) { // NOPMD by Julien on 03/09/15 22:09
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
		final ODXMLDocument stylesDoc = this.sfSpreadSheet.getStyles();
		this.styleHelper.addStyle(stylesDoc, styleName, wrapperCellStyle);
		return true;
	}
}
