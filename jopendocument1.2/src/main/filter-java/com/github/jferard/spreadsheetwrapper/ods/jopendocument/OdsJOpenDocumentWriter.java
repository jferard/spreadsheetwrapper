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
import java.util.ArrayList;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.jdom.Element;
import org.jopendocument.dom.ODXMLDocument;
import org.jopendocument.dom.spreadsheet.Sheet;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.Output;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetReaderCursor;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/
/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * The SpreadSheetWriter for JOpendocument
 *
 */
class OdsJOpenDocumentWriter extends AbstractSpreadsheetDocumentWriter
		implements SpreadsheetDocumentWriter {
	private SpreadsheetWriter createNew(
			/*>>> @UnknownInitialization OdsJOpenDocumentWriterDelegate this, */final Sheet sheet) {
		return new OdsJOpenWriter(this.styleHelper, sheet);
	}

	/** the logger */
	private final Logger logger;

	/** the *internal* workbook */
	private final InitializableDocument initializableDocument;

	/** utility for style */
	private final OdsJOpenStyleHelper styleHelper;

	/**
	 * @param logger
	 *            the logger
	 * @param styleHelper
	 *            the basic style helper
	 * @param initializableDocument
	 *            the *internal* workbook
	 * @param output
	 *            where to write
	 * @param newDocument
	 * @throws SpreadsheetException
	 *             if can't open initializableDocument writer
	 */
	OdsJOpenDocumentWriter(final Logger logger,
			final OdsJOpenStyleHelper styleHelper,
			final InitializableDocument initializableDocument,
			final Output output) throws SpreadsheetException {
		super(logger, output);
		this.styleHelper = styleHelper;
		this.logger = logger;
		this.initializableDocument = initializableDocument;
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
	}

	/** */
	@Override
	public void save() throws SpreadsheetException {
		OutputStream outputStream = null;
		try {
			outputStream = this.output.getStream();
			if (outputStream == null)
				throw new IllegalStateException(String.format(
						"Use saveAs when output file/stream is not specified at opening"));

			this.initializableDocument.save(outputStream);
		} catch (final Exception e) { // NOPMD by Julien on 03/09/15 22:09
			this.logger.log(Level.SEVERE,
					String.format("this.spreadsheetDocument.save(%s) not ok",
							outputStream),
					e);
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyle(final String styleName,
			final WrapperCellStyle wrapperCellStyle) {
		final ODXMLDocument stylesDoc = this.initializableDocument.getStyles();
		this.styleHelper.addStyle(stylesDoc, styleName, wrapperCellStyle);
		return true;
	}

	/** {@inheritDoc} */
	@Override
	public WrapperCellStyle getCellStyle(final String styleName) {
		throw new UnsupportedOperationException();
	}

	/** */
	@Override
	public List<String> getSheetNames() {
		final List<String> sheetNames;
		final int sheetCount = this.getSheetCount();
		sheetNames = new ArrayList<String>(sheetCount);
		for (int s = 0; s < sheetCount; s++)
			sheetNames.add(this.initializableDocument.getSheet(s)
					.getName());
		return sheetNames;
	}

	/**
	 * @param sheetName
	 *            the name of the sheet to find
	 * @throws NoSuchElementException
	 *             if the workbook does not contains any sheet of this name
	 * @return the reader/writer
	 */
	public SpreadsheetWriter getSpreadsheet(final String sheetName)
			throws NoSuchElementException {
		final SpreadsheetWriter spreadsheet;
		if (this.accessor.hasByName(sheetName))
			spreadsheet = this.accessor.getByName(sheetName);
		else
			spreadsheet = this
					.findSpreadsheetAndCreateReaderOrWriter(sheetName);

		return spreadsheet;
	}

	/**
	 * @param index
	 *            index of the sheet
	 * @return the reader/writer
	 */
	public SpreadsheetWriter getSpreadsheet(final int index) {
		final SpreadsheetWriter spreadsheet;
		if (this.accessor.hasByIndex(index))
			spreadsheet = this.accessor.getByIndex(index);
		else {
			if (index < 0 || index >= this.getSheetCount())
				throw new IndexOutOfBoundsException(
						String.format("No sheet at position %d", index));

			final Sheet sheet = this.initializableDocument.getSheet(index);
			spreadsheet = this.createNew(sheet);
			this.accessor.put(sheet.getName(), index, spreadsheet);
		}
		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetWriter addSheetWithCheckedIndex(final int index,
			final String sheetName) {
		Sheet sheet = this.initializableDocument.addSheet(index, sheetName);
		final SpreadsheetWriter spreadsheet = this.createNew(sheet);
		this.accessor.put(sheetName, index, spreadsheet);
		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetWriter findSpreadsheetAndCreateReaderOrWriter(
			final String sheetName) {
		for (int index = 0; index < this.initializableDocument.getSheetCount(); index++) {
			final Sheet sheet = this.initializableDocument.getSheet(index);
			final String name = sheet.getName();
			if (name.equals(sheetName)) {
				final SpreadsheetWriter spreadsheet = this.createNew(sheet);
				this.accessor.put(sheetName, index, spreadsheet);
				return spreadsheet;
			}
		}

		throw new NoSuchElementException(
				String.format("No %s sheet in workbook", sheetName));
	}

	/** {@inheritDoc} */
	@Override
	public int getSheetCount() {
		return this.initializableDocument.getSheetCount();
	}
}
