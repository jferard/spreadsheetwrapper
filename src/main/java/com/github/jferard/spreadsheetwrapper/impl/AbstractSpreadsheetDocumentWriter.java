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
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URL;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;

public abstract class AbstractSpreadsheetDocumentWriter implements
SpreadsheetDocumentWriter {

	/** the logger */
	private final Logger logger;

	/** where to write */
	protected OutputStream outputStream;

	public AbstractSpreadsheetDocumentWriter(final Logger logger,
			final/*@Nullable*/OutputStream outputStream) {
		super();
		this.logger = logger;
		this.outputStream = outputStream;
	}

	@Override
	public void saveAs(final File outputFile) throws SpreadsheetException {
		try {
			final OutputStream outputStream = new FileOutputStream(outputFile);
			this.saveAs(outputStream);
		} catch (final FileNotFoundException e) {
			this.logger.log(Level.SEVERE, String.format(
					"this.spreadsheetDocument.save(%s) not ok",
					this.outputStream), e);
			throw new SpreadsheetException(e);
		}
	}

	@Override
	public void saveAs(final OutputStream outputStream)
			throws SpreadsheetException {
		final OutputStream bkpOutputStream = this.outputStream;

		try {
			this.outputStream = outputStream;
			this.save();
		} catch (final SpreadsheetException e) {
			this.outputStream = bkpOutputStream;
			throw e;
		}
	}

	/** {@inheritDoc} */
	@Override
	public void saveAs(final URL outputURL) throws SpreadsheetException {
		OutputStream outputStream;
		try {
			outputStream = AbstractSpreadsheetDocumentTrait
					.getOutputStream(outputURL);
			this.saveAs(outputStream);
		} catch (final FileNotFoundException e) {
			this.logger.log(Level.SEVERE, String.format(
					"this.spreadsheetDocument.save(%s) not ok",
					this.outputStream), e);
			throw new SpreadsheetException(e);
		} catch (final IOException e) {
			this.logger.log(Level.SEVERE, String.format(
					"this.spreadsheetDocument.save(%s) not ok",
					this.outputStream), e);
			throw new SpreadsheetException(e);
		}
	}

}