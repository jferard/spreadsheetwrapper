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

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URL;
import java.net.URLConnection;
import java.net.UnknownServiceException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.jferard.spreadsheetwrapper.CantInsertElementInSpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;
import com.github.jferard.spreadsheetwrapper.WrapperColor;
import com.github.jferard.spreadsheetwrapper.impl.AbstractSpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.impl.SpreadsheetWriterCursorImpl;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/
/*>>> import org.checkerframework.checker.nullness.qual.RequiresNonNull;*/
/*>>> import org.checkerframework.checker.initialization.qual.UnknownInitialization;*/

/**
 * A wrapper for writing in a workbook
 */
public class XlsPoiDocumentWriter extends AbstractSpreadsheetDocumentWriter
		implements SpreadsheetDocumentWriter {
	/**
	 * A helper, for delegation
	 */
	private final class XlsPoiDocumentWriterTrait extends
			AbstractXlsPoiDocumentTrait<SpreadsheetWriter> {
		/** a map styleName -> internal cell style */
		private final Map<String, CellStyle> cellStyleByName;

		/**
		 * @param workbook
		 *            *internal* workbook
		 * @param dateCellStyle
		 *            the style for all date cells
		 * @param cellStyleByName
		 */
		XlsPoiDocumentWriterTrait(final Workbook workbook,
				final CellStyle dateCellStyle,
				final Map<String, CellStyle> cellStyleByName) {
			super(workbook, dateCellStyle);
			this.cellStyleByName = cellStyleByName;
		}

		/** {@inheritDoc} */
		@Override
		protected SpreadsheetWriter createNew(
				/*>>> @UnknownInitialization XlsPoiDocumentWriterTrait this, */final Sheet sheet) {
			return new XlsPoiWriter(sheet, this.dateCellStyle,
					this.cellStyleByName);
		}
	}

	/**
	 * @param outputURL
	 *            the URL of the file to write
	 * @return the stream on the file
	 * @throws IOException
	 * @throws FileNotFoundException
	 * @deprecated
	 */
	@Deprecated
	private static OutputStream getOutputStream(final URL outputURL)
			throws IOException, FileNotFoundException {
		OutputStream outputStream;
		final URLConnection connection = outputURL.openConnection();
		connection.setDoOutput(true);
		try {
			outputStream = connection.getOutputStream();
		} catch (final UnknownServiceException e) {
			outputStream = new FileOutputStream(outputURL.getPath());
		}
		return outputStream;
	}

	/** a map styleName -> internal cell style */
	private final Map<String, CellStyle> cellStyleByName;
	/** for delegation */
	private final AbstractXlsPoiDocumentTrait<SpreadsheetWriter> documentTrait;

	/** simple logger */
	private final Logger logger;

	/** for delegation */
	private final XlsPoiDocumentReader reader;
	/** for delegation */
	private final XlsPoiStyleUtility styleUtility;
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
	public XlsPoiDocumentWriter(final Logger logger,
			final XlsPoiStyleUtility styleUtility, final Workbook workbook,
			final/*@Nullable*/OutputStream outputStream) {
		super(logger, outputStream);
		this.styleUtility = styleUtility;
		this.reader = new XlsPoiDocumentReader(styleUtility, workbook);
		this.logger = logger;
		this.workbook = workbook;
		final CreationHelper createHelper = this.workbook.getCreationHelper();
		final CellStyle dateCellStyle = this.workbook.createCellStyle();
		this.cellStyleByName = new HashMap<String, CellStyle>();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(
				"yyyy-mm-dd"));
		this.documentTrait = new XlsPoiDocumentWriterTrait(workbook,
				dateCellStyle, this.cellStyleByName);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter addSheet(final int index, final String sheetName)
			throws IndexOutOfBoundsException,
			CantInsertElementInSpreadsheetException {
		return this.documentTrait.addSheet(index, sheetName);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetWriter addSheet(final String sheetName)
			throws CantInsertElementInSpreadsheetException {
		return this.documentTrait.addSheet(sheetName);
	}

	/** {@inheritDoc} */
	@Override
	public void close() {
		this.reader.close();
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public boolean createStyle(final String styleName, final String styleString) {
		final CellStyle cellStyle = this.styleUtility.getCellStyle(
				this.workbook, styleString);
		this.cellStyleByName.put(styleName, cellStyle);
		return true;
	}

	/** {@inheritDoc} */
	@Override
	public com.github.jferard.spreadsheetwrapper.WrapperCellStyle getCellStyle(
			final String styleName) {
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
	@Deprecated
	public String getStyleString(final String styleName) {
		return this.reader.getStyleString(styleName);
	}

	/** */
	@Override
	public void save() throws SpreadsheetException {
		if (this.outputStream == null)
			throw new IllegalStateException(
					String.format("Use saveAs when output file is not specified"));
		try {
			this.workbook.write(this.outputStream);
			this.outputStream.close();
		} catch (final IOException e) {
			this.logger.log(Level.SEVERE, String.format(
					"this.spreadsheetDocument.save(%s) not ok",
					this.outputStream), e);
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public boolean setStyle(
			final String styleName,
			final com.github.jferard.spreadsheetwrapper.WrapperCellStyle wrapperCellStyle) {
		final CellStyle cellStyle = this.workbook.createCellStyle();
		final com.github.jferard.spreadsheetwrapper.WrapperFont wrapperFont = wrapperCellStyle
				.getCellFont();
		if (wrapperFont.isBold()) {
			final Font font = this.workbook.createFont();
			font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			cellStyle.setFont(font);
		}
		final WrapperColor backgroundColor = wrapperCellStyle
				.getBackgroundColor();
		if (backgroundColor != null) {
			final HSSFColor hssfColor = this.styleUtility
					.getHSSFColor(backgroundColor);
			final short index = hssfColor.getIndex();
			cellStyle.setFillForegroundColor(index);
			cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		}
		this.cellStyleByName.put(styleName, cellStyle);
		return true;
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public boolean updateStyle(final String styleName, final String styleString) {
		this.createStyle(styleName, styleString);
		return true;
	}

}
