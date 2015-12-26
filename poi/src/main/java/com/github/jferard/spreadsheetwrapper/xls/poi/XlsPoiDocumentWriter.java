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
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.Output;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetWriterCursorImpl;
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
	/**
	 * A helper, for delegation
	 */
	private final class XlsPoiDocumentWriterDelegate extends
			AbstractXlsPoiDocumentDelegate<SpreadsheetWriter> {

		/**
		 * @param workbook
		 *            *internal* workbook
		 * @param delegateStyleHelper
		 * @param dateCellStyle
		 *            the style for all date cells
		 * @param cellStyleAccessor
		 * @param cellStyleByName
		 */
		XlsPoiDocumentWriterDelegate(final Workbook workbook,
				final XlsPoiStyleHelper styleHelper,
				final CellStyle dateCellStyle) {
			super(workbook, styleHelper, dateCellStyle);
		}

		/** {@inheritDoc} */
		@Override
		/*@RequiresNonNull("delegateStyleHelper")*/
		protected SpreadsheetWriter createNew(
				/*>>> @UnknownInitialization XlsPoiDocumentWriterDelegate this, */final Sheet sheet) {
			return new XlsPoiWriter(sheet, this.delegateStyleHelper,
					this.dateCellStyle);
		}
	}

	/** for delegation */
	private final AbstractXlsPoiDocumentDelegate<SpreadsheetWriter> documentDelegate;

	/** simple logger */
	private final Logger logger;

	/** for delegation */
	private final XlsPoiDocumentReader reader;
	/** for delegation */
	private final XlsPoiStyleHelper styleHelper;
	/** *internal* workbook */
	private final Workbook workbook;

	/**
	 * @param logger
	 *            simple logger
	 * @param workbook
	 *            *internal* workbook
	 * @param outputURL
	 *            where to write
	 */
	public XlsPoiDocumentWriter(final Logger logger, final Workbook workbook,
			final XlsPoiStyleHelper styleHelper, final Output output) {
		super(logger, output);
		this.logger = logger;
		this.workbook = workbook;
		this.styleHelper = styleHelper;
		this.reader = new XlsPoiDocumentReader(workbook, styleHelper);
		final CreationHelper createHelper = this.workbook.getCreationHelper();
		final CellStyle dateCellStyle = this.workbook.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(
				"yyyy-mm-dd"));
		this.documentDelegate = new XlsPoiDocumentWriterDelegate(workbook,
				styleHelper, dateCellStyle);
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
	public/*@Nullable*/WrapperCellStyle getCellStyle(final String styleName) {
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
						String.format("Use saveAs when output file/stream is not specified"));
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
		final CellStyle cellStyle = this.styleHelper.toCellStyle(this.workbook,
				wrapperCellStyle);
		this.styleHelper.putCellStyle(styleName, cellStyle);
		return true;
	}
}
