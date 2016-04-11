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

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.OptionalOutput;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

/*>>>
import org.checkerframework.checker.nullness.qual.Nullable;
import org.checkerframework.checker.initialization.qual.UnknownInitialization;
import org.checkerframework.checker.nullness.qual.RequiresNonNull;
 */

/**
 * A wrapper for writing in a workbook
 */
public class XlsPoiDocumentWriter extends AbstractSpreadsheetDocumentWriter
		implements SpreadsheetDocumentWriter {

		/** simple logger */
		private final Logger logger;

	/** for delegation */
	private final XlsPoiStyleHelper styleHelper;

	/** *internal* workbook */
	private final Workbook workbook;
	private CellStyle dateCellStyle;

	/**
	 * @param logger
	 *            simple logger
	 * @param workbook
	 *            *internal* workbook
	 * @param outputURL
	 *            where to write
	 */
	public XlsPoiDocumentWriter(final Logger logger, final Workbook workbook,
			final XlsPoiStyleHelper styleHelper, final OptionalOutput optionalOutput) {
		super(logger, optionalOutput);
		this.logger = logger;
		this.workbook = workbook;
		this.styleHelper = styleHelper;
		final CreationHelper createHelper = this.workbook.getCreationHelper();
		this.dateCellStyle = this.workbook.createCellStyle();
		this.dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(
				"yyyy-mm-dd"));
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
	public/*@Nullable*/WrapperCellStyle getCellStyle(final String styleName) {
		return this.styleHelper.getWrapperCellStyle(this.workbook, styleName);
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
	public SpreadsheetWriter getSpreadsheet(final int index) {
		final SpreadsheetWriter spreadsheet;
		if (this.accessor.hasByIndex(index))
			spreadsheet = this.accessor.getByIndex(index);
		else {
			if (index < 0 || index >= this.getSheetCount())
				throw new IndexOutOfBoundsException(String.format(
						"No sheet at position %d", index));
		
			final Sheet sheet = this.workbook.getSheetAt(index);
			spreadsheet = this.createNew(sheet);
			this.accessor.put(sheet.getSheetName(), index, spreadsheet);
		}
		return spreadsheet;
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter getSpreadsheet(final String sheetName) {
		final SpreadsheetWriter spreadsheet;
		if (this.accessor.hasByName(sheetName))
			spreadsheet = this.accessor.getByName(sheetName);
		else
			spreadsheet = this
			.findSpreadsheetAndCreateReaderOrWriter(sheetName);
		
		return spreadsheet;
	}
	
	/** */
	@Override
	public void save() throws SpreadsheetException {
		OutputStream outputStream = null;
		try {
			outputStream = this.optionalOutput.getStream();
			if (outputStream == null)
				throw new IllegalStateException(
						String.format("Use saveAs when optionalOutput file/stream is not specified"));
			this.workbook.write(outputStream);
			outputStream.flush();
		} catch (final IOException e) {
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
		return this.styleHelper.setStyle(this.workbook, styleName, wrapperCellStyle);
	}

	@Override
	protected SpreadsheetWriter addSheetWithCheckedIndex(int index,
			String sheetName) throws CantInsertElementInSpreadsheetException {
		Sheet sheet = this.workbook.getSheet(sheetName);
		if (sheet != null)
			throw new IllegalArgumentException(String.format("Sheet %s exists",
					sheetName));

		sheet = this.workbook.createSheet(sheetName);
		this.workbook.setSheetOrder(sheetName, index);
		final SpreadsheetWriter spreadsheet = this.createNew(sheet);
		this.accessor.put(sheetName, index, spreadsheet);
		return spreadsheet;
	}

	/*@RequiresNonNull("delegateStyleHelper")*/
	protected SpreadsheetWriter createNew(
			/*>>> @UnknownInitialization XlsPoiDocumentWriterDelegate this, */final Sheet sheet) {
		return new XlsPoiWriter(sheet, this.styleHelper,
				this.dateCellStyle);
	}

	@Override
	protected SpreadsheetWriter findSpreadsheetAndCreateReaderOrWriter(
			String sheetName) throws NoSuchElementException {
		final Sheet sheet = this.workbook.getSheet(sheetName);
		if (sheet == null)
			throw new NoSuchElementException(String.format(
					"No %s sheet in workbook", sheetName));

		final int index = this.workbook.getSheetIndex(sheet);
		final SpreadsheetWriter spreadsheet = this.createNew(sheet);
		this.accessor.put(sheetName, index, spreadsheet);
		return spreadsheet;
	}
}
