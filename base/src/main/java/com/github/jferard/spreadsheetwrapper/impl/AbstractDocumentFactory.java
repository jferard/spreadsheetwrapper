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
import java.io.InputStream;
import java.io.OutputStream;
import java.util.logging.Logger;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * @param <R>
 *            internal workbook
 */
public abstract class AbstractDocumentFactory<R> extends
AbstractBasicDocumentFactory implements SpreadsheetDocumentFactory {
	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create(
			final/*@Nullable*/OutputStream outputStream)
					throws SpreadsheetException {
		final R document = this.newSpreadsheetDocument(outputStream);
		return this.createWriter(Stateful.createNew(document), new Output(
				outputStream));
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentReader openForRead(final InputStream inputStream)
			throws SpreadsheetException {
		final R document = this.loadSpreadsheetDocument(inputStream);
		return this.createReader(Stateful.createInitialized(document));
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter openForWrite(final File inputFile,
			final/*@Nullable*/File outputFile) throws SpreadsheetException {
		InputStream inputStream;
		try {
			inputStream = new FileInputStream(inputFile);
			final R document = this.loadSpreadsheetDocument(inputStream);
			return this.createWriter(Stateful.createInitialized(document),
					new Output(outputFile));
		} catch (final FileNotFoundException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter openForWrite(
			final InputStream inputStream,
			final/*@Nullable*/OutputStream outputStream)
					throws SpreadsheetException {
		final R document = this.loadSpreadsheetDocument(inputStream);
		return this.createWriter(Stateful.createInitialized(document),
				new Output(outputStream));
	}

	/**
	 * create a reader from a *internal* workbook
	 *
	 * @param stateful
	 *            the internal workbook
	 * @return the value reader
	 * @throws SpreadsheetException
	 */
	protected abstract SpreadsheetDocumentReader createReader(
			Stateful<R> stateful) throws SpreadsheetException;

	/**
	 * create a writer from a *internal* workbook
	 *
	 * @param stateful
	 *            the internal workbook
	 * @return the value writer
	 * @throws SpreadsheetException
	 */
	protected abstract SpreadsheetDocumentWriter createWriter(
			Stateful<R> stateful, Output output) throws SpreadsheetException;

	/**
	 * @param inputStream
	 *            the source of the value
	 * @return the *internal* value
	 * @throws SpreadsheetException
	 */
	protected abstract R loadSpreadsheetDocument(InputStream inputStream)
			throws SpreadsheetException;

	/**
	 * @param outputStream
	 *            the destination of the value
	 * @return the *internal* value
	 * @throws SpreadsheetException
	 */
	protected abstract R newSpreadsheetDocument(
			/*@Nullable*/OutputStream outputStream) throws SpreadsheetException;

}