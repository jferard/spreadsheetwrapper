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
package com.github.jferard.spreadsheetwrapper.ods.simpleodf;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.logging.Logger;

import org.odftoolkit.simple.SpreadsheetDocument;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.impl.AbstractDocumentFactory;
import com.github.jferard.spreadsheetwrapper.impl.OptionalOutput;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;
import com.github.jferard.spreadsheetwrapper.ods.apache.OdsOdfdomStyleHelper;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class OdsSimpleodfDocumentFactory
		extends AbstractDocumentFactory<InitializableDocument>
		implements SpreadsheetDocumentFactory {
	/**
	 * Static method used by the manager for creating a factory
	 *
	 * @param logger
	 *            the logger
	 * @return the factory
	 */
	public static SpreadsheetDocumentFactory create(final Logger logger) {
		return new OdsSimpleodfDocumentFactory(logger,
				new OdsOdfdomStyleHelper());
	}

	/** simple logger */
	private final Logger logger;

	/** style helper */
	private final OdsOdfdomStyleHelper styleHelper;

	/**
	 * @param logger
	 *            the logger
	 * @param styleHelper
	 *            the style helper
	 */
	public OdsSimpleodfDocumentFactory(final Logger logger,
			final OdsOdfdomStyleHelper styleHelper) {
		super(OdsConstants.EXTENSION);
		this.logger = logger;
		this.styleHelper = styleHelper;
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetDocumentWriter createWriter(
			final InitializableDocument initializableDocument,
			final OptionalOutput optionalOutput) throws SpreadsheetException {
		return new OdsSimpleodfDocumentWriter(this.logger, this.styleHelper,
				initializableDocument, optionalOutput);
	}

	/** {@inheritDoc} */
	@Override
	protected InitializableDocument loadSpreadsheetDocument(
			final InputStream inputStream) throws SpreadsheetException {
		try {
			return InitializableDocument.createInitialized(
					SpreadsheetDocument.loadDocument(inputStream));
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		} catch (final Exception e) { // NOPMD by Julien on 28/11/15 16:11
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	protected InitializableDocument newSpreadsheetDocument(
			final/*@Nullable*/OutputStream outputStream)
			throws SpreadsheetException {
		try {
			return InitializableDocument.createUninitialized(
					SpreadsheetDocument.newSpreadsheetDocument());
		} catch (final Exception e) { // NOPMD by Julien on 28/11/15 16:11
			throw new SpreadsheetException(e);
		}
	}
}
