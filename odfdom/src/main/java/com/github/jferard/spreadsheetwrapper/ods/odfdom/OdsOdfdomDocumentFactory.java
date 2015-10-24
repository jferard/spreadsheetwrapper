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

package com.github.jferard.spreadsheetwrapper.ods.odfdom;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.logging.Logger;

import org.odftoolkit.odfdom.doc.OdfSpreadsheetDocument;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.impl.AbstractDocumentFactory;
import com.github.jferard.spreadsheetwrapper.impl.Output;
import com.github.jferard.spreadsheetwrapper.impl.Stateful;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class OdsOdfdomDocumentFactory extends
		AbstractDocumentFactory<OdfSpreadsheetDocument> implements
		SpreadsheetDocumentFactory {
	public static SpreadsheetDocumentFactory create(final Logger logger) {
		return new OdsOdfdomDocumentFactory(logger, new OdsOdfdomStyleUtility());
	}

	/** simple logger */
	private final Logger logger;

	private final OdsOdfdomStyleUtility styleUtility;

	/**
	 * @param logger
	 *            simple logger
	 */
	public OdsOdfdomDocumentFactory(final Logger logger,
			final OdsOdfdomStyleUtility styleUtility) {
		super(logger);
		this.logger = logger;
		this.styleUtility = styleUtility;
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetDocumentReader createReader(
			final Stateful<OdfSpreadsheetDocument> sfDocument)
			throws SpreadsheetException {
		return new OdsOdfdomDocumentReader(this.styleUtility,
				sfDocument.getObject());
	}

	/**
	 * {@inheritDoc}
	 *
	 * @throws SpreadsheetException
	 */
	@Override
	protected SpreadsheetDocumentWriter createWriter(
			final Stateful<OdfSpreadsheetDocument> sfDocument,
			final Output output)
			throws SpreadsheetException {
		return new OdsOdfdomDocumentWriter(this.logger, this.styleUtility,
				sfDocument.getObject(), output);
	}

	@Override
	protected OdfSpreadsheetDocument loadSpreadsheetDocument(
			final InputStream inputStream) throws SpreadsheetException {
		OdfSpreadsheetDocument document;
		try {
			document = OdfSpreadsheetDocument.loadDocument(inputStream);
		} catch (final Exception e) { // NOPMD by Julien on 03/09/15 22:04
			throw new SpreadsheetException(e);
		}
		return document;
	}

	@Override
	protected OdfSpreadsheetDocument newSpreadsheetDocument(
			final/*@Nullable*/OutputStream outputStream)
			throws SpreadsheetException {
		OdfSpreadsheetDocument document;
		try {
			document = OdfSpreadsheetDocument.newSpreadsheetDocument();
			document.getTableList().get(0).remove(); // a sheet is already
			// present
		} catch (final Exception e) { // NOPMD by Julien on 03/09/15 22:04
			throw new SpreadsheetException(e);
		}
		return document;
	}
}
