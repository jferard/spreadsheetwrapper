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

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.simpleods.OdsFile;
import org.simpleods.SimpleOdsException;
import org.simpleods.Table;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.OptionalOutput;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

/**
 */
public class OdsSimpleodsDocumentWriter extends
		AbstractSpreadsheetDocumentWriter implements SpreadsheetDocumentWriter {
	
	/** *internal* workbook */
	private final OdsFile file;

	/** logger */
	private final Logger logger;

	private OdsSimpleodsStyleHelper styleHelper;

	/**
	 * @param logger
	 *            simple logger
	 * @param styleHelper 
	 * @param file
	 *            *internal* value
	 * @param outputURL
	 *            where to write
	 */
	public OdsSimpleodsDocumentWriter(final Logger logger, OdsSimpleodsStyleHelper styleHelper, final OdsFile file,
			final OptionalOutput optionalOutput) {
		super(logger, optionalOutput);
		this.logger = logger;
		this.styleHelper = styleHelper;
		this.file = file;
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		try {
			this.optionalOutput.close();
		} catch (final IOException e) {
			final String message = e.getMessage();
			this.logger.log(Level.SEVERE, message == null ? "" : message, e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public WrapperCellStyle getCellStyle(final String styleName) {
		throw new UnsupportedOperationException();
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

	/**
	 * @param index
	 *            index of the sheet
	 * @return the reader/writer
	 */
	@Override
	public SpreadsheetWriter getSpreadsheet(final int index) {
		final SpreadsheetWriter spreadsheet;
		if (this.accessor.hasByIndex(index))
			spreadsheet = this.accessor.getByIndex(index);
		else {
			final int count = this.getSheetCount();
			if (index < 0 || index >= count)
				throw new NoSuchElementException(String.format(
						"No sheet at position %d", index));

			try {
				final String name = this.file.getTableName(index);
				final Table table = this.getTable(index);
				spreadsheet = this.createNew(table);
				this.accessor.put(name, index, spreadsheet);
			} catch (final SimpleOdsException e) {
				throw new AssertionError(String.format(
						"Can't reach sheet at position %d", index), e);
			}
		}
		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	public void save() throws SpreadsheetException {
		final OutputStream outputStream = this.optionalOutput.getStream();
		if (outputStream == null)
			throw new IllegalStateException(
					String.format("Use saveAs when optionalOutput file/stream is not specified"));
		if (!this.file.save(outputStream))
			throw new SpreadsheetException(String.format(
					"this.spreadsheetDocument.save(%s) not ok", outputStream));
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyle(final String styleName,
			final WrapperCellStyle wrapperCellStyle) {
		return this.styleHelper.setStyle(this.file, styleName, wrapperCellStyle);
	}

	private SpreadsheetWriter createNew(final Table table) {
		return new OdsSimpleodsWriter(this.styleHelper, table);
	}

	private Table getTable(
			/*>>> @UnknownInitialization AbstractOdsSimpleodsDocumentDelegate<T> this, */final int index) {
		if (this.file == null)
			throw new IllegalStateException();

		return (Table) this.file.getContent().getTableQueue().get(index);
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetWriter addSheetWithCheckedIndex(final int index, final String sheetName)
			throws CantInsertElementInSpreadsheetException {
		if (index != this.getSheetCount())
			throw new UnsupportedOperationException();

		final int indexForSheetName = this.file.getTableNumber(sheetName);
		if (indexForSheetName != -1)
			throw new IllegalArgumentException(String.format("Sheet %s exists",
					sheetName));

		// at the end
		final SpreadsheetWriter spreadsheet;
		try {
			this.file.addTable(sheetName);
			final Table table = this.getTable(index);
			spreadsheet = this.createNew(table);
			this.accessor.put(sheetName, index, spreadsheet);
		} catch (final SimpleOdsException e) {
			throw new CantInsertElementInSpreadsheetException(e);
		}
		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetWriter findSpreadsheetAndCreateReaderOrWriter(final String sheetName) {
		final int index = this.file.getTableNumber(sheetName);
		if (index == -1)
			throw new NoSuchElementException(String.format(
					"No %s sheet in workbook", sheetName));

		final Table table = this.getTable(index);
		final SpreadsheetWriter spreadsheet = this.createNew(table);
		this.accessor.put(sheetName, index, spreadsheet);
		return spreadsheet;
	}
}
