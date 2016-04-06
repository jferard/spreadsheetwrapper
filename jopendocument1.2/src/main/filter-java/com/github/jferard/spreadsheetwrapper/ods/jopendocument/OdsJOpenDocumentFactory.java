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

package com.github.jferard.spreadsheetwrapper.ods.${jopendocument.pkg};

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.logging.Logger;

import org.jopendocument.dom.ODPackage;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.Output;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.impl.AbstractDocumentFactory;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class OdsJOpenDocumentFactory extends
AbstractDocumentFactory<InitializableDocument> implements
SpreadsheetDocumentFactory {
	/**
	 * static method for the manager
	 *
	 * @param logger
	 *            the logger
	 * @return a factory
	 */
	public static SpreadsheetDocumentFactory create(final Logger logger) {
		return new OdsJOpenDocumentFactory(logger, new OdsJOpenStyleHelper());
	}

	/** simple logger */
	private final Logger logger;

	/** utility for styles */
	private final OdsJOpenStyleHelper styleHelper;

	/**
	 * @param logger
	 *            simple logger
	 * @param styleHelper
	 *            utility for styles
	 */
	OdsJOpenDocumentFactory(final Logger logger,
			final OdsJOpenStyleHelper styleHelper) {
		super(OdsConstants.EXTENSION);
		this.logger = logger;
		this.styleHelper = styleHelper;
	}

	/**
	 * {@inheritDoc}
	 *
	 * @throws SpreadsheetException
	 */
	@Override
	protected SpreadsheetDocumentWriter createWriter(
			final InitializableDocument initializableDocument, final Output output)
					throws SpreadsheetException {
		return new OdsJOpenDocumentWriter(this.logger, this.styleHelper,
				initializableDocument, output);
	}

	/** {@inheritDoc} */
	@Override
	protected InitializableDocument loadSpreadsheetDocument(final InputStream inputStream)
			throws SpreadsheetException {
		try {
			final ODPackage odPackage = new ODPackage(inputStream);
			return InitializableDocument.createInitialized(${jopendocument.util.cls}.getSpreadSheet(odPackage));
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	protected InitializableDocument newSpreadsheetDocument(
			final/*@Nullable*/OutputStream outputStream)
					throws SpreadsheetException {
		return InitializableDocument.createUninitialized(${jopendocument.util.cls}.newSpreadsheetDocument());
	}

}
