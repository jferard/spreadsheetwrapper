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
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.util.Locale;
import java.util.logging.Logger;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.WritableWorkbook;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.impl.AbstractBasicDocumentFactory;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class XlsJxlDocumentFactory extends AbstractBasicDocumentFactory
implements SpreadsheetDocumentFactory {
	public static SpreadsheetDocumentFactory create(final Logger logger) {
		return new XlsJxlDocumentFactory(logger, new XlsJxlStyleUtility());
	}

	/** simple logger */
	private final Logger logger;

	private final XlsJxlStyleUtility styleUtility;

	/**
	 * @param logger
	 *            simple logger
	 */
	public XlsJxlDocumentFactory(final Logger logger,
			final XlsJxlStyleUtility styleUtility) {
		this.logger = logger;
		this.styleUtility = styleUtility;
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

		try {
			final WritableWorkbook writableWorkbook = Workbook.createWorkbook(
					outputStream, this.getWriteSettings());

			return new XlsJxlDocumentWriter(this.logger, this.styleUtility,
					writableWorkbook);
		} catch (final FileNotFoundException e) {
			throw new SpreadsheetException(e);
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentReader openForRead(final InputStream inputStream)
			throws SpreadsheetException {
		try {
			final Workbook workbook = Workbook.getWorkbook(inputStream,
					this.getReadSettings());
			return new XlsJxlDocumentReader(workbook);
		} catch (final FileNotFoundException e) {
			throw new SpreadsheetException(e);
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		} catch (final BiffException e) {
			throw new SpreadsheetException(e);
		}
	}

	@Override
	public SpreadsheetDocumentWriter openForWrite(final File inputFile,
			final File outputFile) throws SpreadsheetException {
		if (outputFile == null)
			throw new SpreadsheetException("Specify an output file");

		try {
			final Workbook workbook = Workbook.getWorkbook(inputFile,
					this.getReadSettings());
			final WritableWorkbook writableWorkbook = Workbook.createWorkbook(
					outputFile, workbook, this.getWriteSettings());
			return new XlsJxlDocumentWriter(this.logger, this.styleUtility,
					writableWorkbook);
		} catch (final FileNotFoundException e) {
			throw new SpreadsheetException(e);
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		} catch (final BiffException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter openForWrite(final InputStream inputStream) {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter openForWrite(
			final InputStream inputStream, final OutputStream outputStream)
					throws SpreadsheetException {
		if (outputStream == null)
			throw new SpreadsheetException("Specify an output stream");

		try {
			final Workbook workbook = Workbook.getWorkbook(inputStream,
					this.getReadSettings());
			final WritableWorkbook writableWorkbook = Workbook.createWorkbook(
					outputStream, workbook, this.getWriteSettings());
			return new XlsJxlDocumentWriter(this.logger, this.styleUtility,
					writableWorkbook);
		} catch (final FileNotFoundException e) {
			throw new SpreadsheetException(e);
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		} catch (final BiffException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public SpreadsheetDocumentWriter openForWrite(final URL inputURL)
			throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}

	private WorkbookSettings getReadSettings() {
		final WorkbookSettings settings = new WorkbookSettings();
		settings.setLocale(Locale.US);
		settings.setEncoding("windows-1252");
		return settings;
	}

	private WorkbookSettings getWriteSettings() {
		final WorkbookSettings settings = this.getReadSettings();
		settings.setWriteAccess("spreadsheetwrapper");
		return settings;
	}
}
