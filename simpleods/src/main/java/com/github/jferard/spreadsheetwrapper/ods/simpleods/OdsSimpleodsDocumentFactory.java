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
import java.io.InputStream;
import java.io.OutputStream;
import java.util.logging.Logger;

import org.simpleods.OdsFile;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.impl.AbstractBasicDocumentFactory;
import com.github.jferard.spreadsheetwrapper.impl.OptionalOutput;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class OdsSimpleodsDocumentFactory extends AbstractBasicDocumentFactory
implements SpreadsheetDocumentFactory {
	/**
	 * static fuction for the manager
	 *
	 * @param logger
	 *            the logger
	 * @return the factory
	 */
	public static SpreadsheetDocumentFactory create(final Logger logger) {
		return new OdsSimpleodsDocumentFactory(logger);
	}

	/** the logger */
	private final Logger logger;

	/**
	 * @param logger
	 *            the logger
	 */
	public OdsSimpleodsDocumentFactory(final Logger logger) {
		super(OdsConstants.EXTENSION);
		this.logger = logger;
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create() throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create(final/*@Nullable*/File file)
			throws SpreadsheetException {
		if (file == null)
			return this.create();

		final OdsFile odsFile = new OdsFile(file.getPath());
		final OptionalOutput optionalOutput = OptionalOutput.fromFile(file);
		return new OdsSimpleodsDocumentWriter(this.logger, odsFile, optionalOutput);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create(
			final/*@Nullable*/OutputStream outputStream)
					throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentReader openForRead(final InputStream inputStream)
			throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter openForWrite(final File inputFile,
			final File outputFile) throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter openForWrite(
			final InputStream inputStream,
			final/*@Nullable*/OutputStream outputStream)
					throws SpreadsheetException {
		throw new UnsupportedOperationException();
	}
}
