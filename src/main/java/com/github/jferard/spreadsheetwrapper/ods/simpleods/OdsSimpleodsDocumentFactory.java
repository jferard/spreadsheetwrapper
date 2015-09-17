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

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.logging.Logger;

import org.simpleods.OdsFile;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.impl.AbstractBasicDocumentFactory;

/*>>> import checkers.nullness.quals.Nullable;*/

public class OdsSimpleodsDocumentFactory extends AbstractBasicDocumentFactory
		implements SpreadsheetDocumentFactory {
	/** the logger */
	private final Logger logger;

	/**
	 * @param logger
	 *            the logger
	 */
	public OdsSimpleodsDocumentFactory(final Logger logger) {
		this.logger = logger;
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create() throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create(final File f)
			throws SpreadsheetException {
		final OdsFile file = new OdsFile(f.getPath());
		OutputStream outputStream;
		try {
			outputStream = new FileOutputStream(f);
			return new OdsSimpleodsDocumentWriter(this.logger, file,
					outputStream);
		} catch (final FileNotFoundException e) {
			throw new SpreadsheetException(e);
		}
	}

	@Override
	public SpreadsheetDocumentWriter create(final OutputStream outputStream)
			throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create(final URL outputURL)
			throws SpreadsheetException {
		try {
			final File f = new File(outputURL.toURI());
			return this.create(f);
		} catch (final IllegalArgumentException e) { // not a file scheme
			throw new SpreadsheetException(e);
		} catch (final URISyntaxException e) {
			throw new SpreadsheetException(e);
		}
	}

	@Override
	public SpreadsheetDocumentReader openForRead(final InputStream inputStream)
			throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}

	@Override
	public SpreadsheetDocumentWriter openForWrite(final File inputFile,
			final File outputFile) throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}

	@Override
	public SpreadsheetDocumentWriter openForWrite(
			final InputStream inputStream, final OutputStream outputStream)
			throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}

	@Override
	public SpreadsheetDocumentWriter openForWrite(final URL inputURL,
			final URL outputURL) throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}
}
