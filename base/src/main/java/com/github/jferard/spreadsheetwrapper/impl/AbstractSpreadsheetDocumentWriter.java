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
package com.github.jferard.spreadsheetwrapper.impl;

import java.io.File;
import java.io.OutputStream;
import java.util.NoSuchElementException;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.github.jferard.spreadsheetwrapper.Accessor;
import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;

/*>>> import org.checkerframework.checker.nullness.qual.MonotonicNonNull;*/

public abstract class AbstractSpreadsheetDocumentWriter implements
SpreadsheetDocumentWriter {

	/** the optionalOutput */
	private/*@MonotonicNonNull*/OptionalOutput bkpOutput;

	/** the logger */
	private final Logger logger;

	/** where to write */
	protected OptionalOutput optionalOutput;

	/**
	 * An accessor on readers/writers, by name and index.
	 */
	protected final Accessor<SpreadsheetWriter> accessor;

	/**
	 * @param logger
	 *            the loggier
	 * @param outputStream
	 *            where to write
	 */
	public AbstractSpreadsheetDocumentWriter(final Logger logger,
			final OptionalOutput optionalOutput) {
		super();
		this.logger = logger;
		this.optionalOutput = optionalOutput;
		this.accessor = new Accessor<SpreadsheetWriter>();
	}

	/**
	 * Adds a sheet to the value
	 *
	 * @param index
	 *            (0..n) the sheet is inserted before this index
	 * @param sheetName
	 *            the namee of the sheet
	 * @throws IndexOutOfBoundsException
	 *             if index < 0 or index > sheetCount
	 * @return the reader/writer
	 * @throws CantInsertElementInSpreadsheetException
	 */
	@Override
	public SpreadsheetWriter addSheet(final int index, final String sheetName)
			throws IndexOutOfBoundsException,
			CantInsertElementInSpreadsheetException {
		final int size = this.getSheetCount();
		if (index < 0 || index > size) // index == size is ok
			throw new IndexOutOfBoundsException();

		return this.addSheetWithCheckedIndex(index, sheetName);
	}
	
	/**
	 * Adds a sheet at the end of the workbook
	 *
	 * @param sheetName
	 *            the name of the sheet
	 * @return the reader/writer
	 * @throws CantInsertElementInSpreadsheetException
	 */
	@Override
	public SpreadsheetWriter addSheet(final String sheetName)
			throws CantInsertElementInSpreadsheetException {
		return this.addSheetWithCheckedIndex(this.getSheetCount(), sheetName);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriterCursor getNewCursorByIndex(final int index) {
		return new SpreadsheetWriterCursorImpl(this.getSpreadsheet(index));
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriterCursor getNewCursorByName(final String sheetName) {
		return new SpreadsheetWriterCursorImpl(this.getSpreadsheet(sheetName));
	}

	/**
	 * @param sheetName
	 *            the name of the sheet to find
	 * @throws NoSuchElementException
	 *             if the workbook does not contains any sheet of this name
	 * @return the reader/writer
	 */
	@Override
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
	
	/** {@inheritDoc} */
	@Override
	public void saveAs(final File outputFile) throws SpreadsheetException {
		this.bkpOutput = this.optionalOutput;
		try {
			this.optionalOutput = OptionalOutput.fromFile(outputFile);
			this.save();
		} catch (final SpreadsheetException e) {
			this.optionalOutput = this.bkpOutput;
			this.logger.log(Level.SEVERE, String.format(
					"this.spreadsheetDocument.save(%s) not ok", outputFile), e);
			throw e;
		}
	}

	/** {@inheritDoc} */
	@Override
	public void saveAs(final OutputStream outputStream)
			throws SpreadsheetException {
		this.bkpOutput = this.optionalOutput;
		try {
			this.optionalOutput = OptionalOutput.fromStream(outputStream);
			this.save();
		} catch (final SpreadsheetException e) {
			this.optionalOutput = this.bkpOutput;
			throw e;
		}
	}

	

	/**
	 * @param index
	 *            the index of the new sheet (0..sheetCount)
	 * @param sheetName
	 *            the name of the sheet
	 * @return the reader/writer on thee sheet, table, ... inserted
	 * @throws CantInsertElementInSpreadsheetException
	 *             if cant inseert the spreadsheet in the workbook
	 */
	protected abstract SpreadsheetWriter addSheetWithCheckedIndex(int index, String sheetName)
			throws CantInsertElementInSpreadsheetException;

	/**
	 * @param sheetName
	 *            the sheet, table, ... name
	 * @throws NoSuchElementException
	 *             if the workbook does not contains any sheet of this name
	 * @return the reader/writer
	 */
	protected abstract SpreadsheetWriter findSpreadsheetAndCreateReaderOrWriter(String sheetName)
			throws NoSuchElementException;
}