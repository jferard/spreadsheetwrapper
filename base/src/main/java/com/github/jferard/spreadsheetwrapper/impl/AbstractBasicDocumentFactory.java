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
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public abstract class AbstractBasicDocumentFactory implements
		SpreadsheetDocumentFactory {

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create() throws SpreadsheetException {
		return this.create((OutputStream) null);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create(final/*@Nullable*/File outputFile)
			throws SpreadsheetException {
		if (outputFile == null)
			return this.create();

		try {
			final FileOutputStream fileOutputStream = new FileOutputStream(
					outputFile);
			return this.create(fileOutputStream);
		} catch (final FileNotFoundException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public SpreadsheetDocumentWriter create(final/*@Nullable*/URL outputURL)
			throws SpreadsheetException {
		if (outputURL == null)
			return this.create();

		try {
			final OutputStream outputStream = AbstractSpreadsheetDocumentTrait
					.getOutputStream(outputURL);
			return this.create(outputStream);
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentReader openForRead(final File inputFile)
			throws SpreadsheetException {
		try {
			return this.openForRead(new FileInputStream(inputFile));
		} catch (final FileNotFoundException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public SpreadsheetDocumentReader openForRead(final URL inputURL)
			throws SpreadsheetException {
		try {
			final InputStream inputStream = inputURL.openStream();
			return this.openForRead(inputStream);
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter openForWrite(final File inputFile)
			throws SpreadsheetException {
		return this.openForWrite(inputFile, inputFile);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter openForWrite(final File inputFile,
			final File outputFile) throws SpreadsheetException {
		try {
			return this.openForWrite(new FileInputStream(inputFile),
					new FileOutputStream(outputFile));
		} catch (final FileNotFoundException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter openForWrite(final InputStream inputStream)
			throws SpreadsheetException {
		return this.openForWrite(inputStream, (OutputStream) null);
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public SpreadsheetDocumentWriter openForWrite(final URL inputURL)
			throws SpreadsheetException {
		return this.openForWrite(inputURL, inputURL);
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public SpreadsheetDocumentWriter openForWrite(final URL inputURL,
			final URL outputURL) throws SpreadsheetException {
		try {
			final InputStream inputStream = inputURL.openStream();
			final OutputStream outputStream = AbstractSpreadsheetDocumentTrait
					.getOutputStream(outputURL);
			return this.openForWrite(inputStream, outputStream);
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}

}