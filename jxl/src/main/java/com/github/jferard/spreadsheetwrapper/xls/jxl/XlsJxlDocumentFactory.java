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
package com.github.jferard.spreadsheetwrapper.xls.jxl;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Locale;
import java.util.logging.Logger;

import jxl.WorkbookSettings;
import jxl.write.WritableCellFormat;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.impl.AbstractBasicDocumentFactory;
import com.github.jferard.spreadsheetwrapper.style.CellStyleAccessor;
import com.github.jferard.spreadsheetwrapper.xls.XlsConstants;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class XlsJxlDocumentFactory extends AbstractBasicDocumentFactory
		implements SpreadsheetDocumentFactory {
	/**
	 * Called by the manager
	 *
	 * @param logger
	 * @return a factory
	 */
	public static SpreadsheetDocumentFactory create(final Logger logger) {
		final XlsJxlStyleColorHelper colorHelper = new XlsJxlStyleColorHelper();
		return new XlsJxlDocumentFactory(logger, new XlsJxlStyleHelper(
				new CellStyleAccessor<WritableCellFormat>(), colorHelper,
				new XlsJxlStyleFontHelper(colorHelper),
				new XlsJxlStyleBorderHelper(colorHelper)));
	}

	/** simple logger */
	private final Logger logger;

	/** helper for style */
	private final XlsJxlStyleHelper styleHelper;

	/**
	 * @param logger
	 *            simple logger
	 */
	XlsJxlDocumentFactory(final Logger logger,
			final XlsJxlStyleHelper styleHelper) {
		super(XlsConstants.EXTENSION1);
		this.logger = logger;
		this.styleHelper = styleHelper;
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create() throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create(
			final/*@Nullable*/OutputStream outputStream)
			throws SpreadsheetException {
		if (outputStream == null)
			return this.create();

		final JxlWorkbook jxlWorkbook = new JxlWritableWorkbook(outputStream);
		return new XlsJxlDocumentWriter(this.logger, this.styleHelper,
				jxlWorkbook);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentReader openForRead(final InputStream inputStream)
			throws SpreadsheetException {
		JxlWorkbook jxlWorkbook = new JxlReadableWorkbook(inputStream);   
		return new XlsJxlDocumentWriter(this.logger, this.styleHelper, jxlWorkbook);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter openForWrite(final File inputFile,
			final File outputFile) throws SpreadsheetException {
		final JxlWorkbook jxlWorkbook = new JxlWritableWorkbook(inputFile,
				outputFile);
		return new XlsJxlDocumentWriter(this.logger, this.styleHelper,
				jxlWorkbook);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter openForWrite(
			final InputStream inputStream,
			final/*@Nullable*/OutputStream outputStream)
			throws SpreadsheetException {
		final JxlWorkbook jxlWorkbook = new JxlWritableWorkbook(inputStream,
				outputStream);
		return new XlsJxlDocumentWriter(this.logger, this.styleHelper,
				jxlWorkbook);
	}

	private static WorkbookSettings getWriteSettings() {
		final WorkbookSettings settings = getReadSettings();
		settings.setWriteAccess("spreadsheetwrapper");
		return settings;
	}

	private static WorkbookSettings getReadSettings() {
		final WorkbookSettings settings = new WorkbookSettings();
		settings.setLocale(Locale.US);
		settings.setEncoding("windows-1252");
		return settings;
	}

}
