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

/*>>> import checkers.nullness.quals.Nullable;*/

public class XlsJxlDocumentFactory extends AbstractBasicDocumentFactory
		implements SpreadsheetDocumentFactory {
	/** simple logger */
	private final Logger logger;

	/**
	 * @param logger
	 *            simple logger
	 */
	public XlsJxlDocumentFactory(final Logger logger) {
		this.logger = logger;
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create() throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create(final OutputStream outputStream)
			throws SpreadsheetException {
		try {
			final WritableWorkbook writableWorkbook = Workbook
					.createWorkbook(outputStream);
			return new XlsJxlDocumentWriter(this.logger, writableWorkbook);
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
			final Workbook workbook = Workbook.getWorkbook(inputStream);
			return new XlsJxlDocumentReader(workbook);
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
	public SpreadsheetDocumentWriter openForWrite(
			final InputStream inputStream, final OutputStream outputStream)
			throws SpreadsheetException {
		try {
			final Workbook workbook = Workbook.getWorkbook(inputStream);
			final WorkbookSettings settings = new WorkbookSettings();
			settings.setLocale(Locale.US);
			settings.setWriteAccess("spreadsheetwrapper");
			final WritableWorkbook writableWorkbook = Workbook.createWorkbook(
					outputStream, workbook, settings);
			return new XlsJxlDocumentWriter(this.logger, writableWorkbook);
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
	public SpreadsheetDocumentWriter openForWrite(final URL inputURL)
			throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}
}
