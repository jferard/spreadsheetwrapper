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
import java.io.OutputStream;
import java.net.URL;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;

/*>>> import org.checkerframework.checker.nullness.qual.MonotonicNonNull;*/

public abstract class AbstractSpreadsheetDocumentWriter implements
		SpreadsheetDocumentWriter {

	private/*@MonotonicNonNull*/Output bkpOutput;

	/** the logger */
	private final Logger logger;

	/** where to write */
	protected Output output;

	/**
	 * @param logger
	 *            the loggier
	 * @param outputStream
	 *            where to write
	 */
	public AbstractSpreadsheetDocumentWriter(final Logger logger,
			final Output output) {
		super();
		this.logger = logger;
		this.output = output;
	}

	/** {@inheritDoc} */
	@Override
	public void saveAs(final File outputFile) throws SpreadsheetException {
		this.bkpOutput = this.output;
		try {
			this.output = new Output(outputFile);
			this.save();
		} catch (final SpreadsheetException e) {
			this.output = this.bkpOutput;
			this.logger.log(Level.SEVERE, String.format(
					"this.spreadsheetDocument.save(%s) not ok", outputFile), e);
			throw e;
		}
	}

	/** {@inheritDoc} */
	@Override
	public void saveAs(final OutputStream outputStream)
			throws SpreadsheetException {
		this.bkpOutput = this.output;
		try {
			this.output = new Output(outputStream);
			this.save();
		} catch (final SpreadsheetException e) {
			this.output = this.bkpOutput;
			throw e;
		}
	}

	/** {@inheritDoc} */
	@Override
	@Deprecated
	public void saveAs(final URL outputURL) throws SpreadsheetException {
		this.bkpOutput = this.output;
		try {
			this.output = new Output(outputURL);
			this.save();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.SEVERE, String.format(
					"this.spreadsheetDocument.save(%s) not ok", outputURL), e);
			throw e;
		}
	}
}