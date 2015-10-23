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
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.impl.AbstractDocumentFactory;
import com.github.jferard.spreadsheetwrapper.impl.Stateful;
import com.github.jferard.spreadsheetwrapper.ods.odfdom.OdsOdfdomStyleUtility;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class OdsSimpleodfDocumentFactory extends
AbstractDocumentFactory<SpreadsheetDocument> implements
SpreadsheetDocumentFactory {
	public static SpreadsheetDocumentFactory create(final Logger logger) {
		return new OdsSimpleodfDocumentFactory(logger,
				new OdsOdfdomStyleUtility());
	}

	private final Logger logger;

	private final OdsOdfdomStyleUtility styleUtility;

	public OdsSimpleodfDocumentFactory(final Logger logger,
			final OdsOdfdomStyleUtility styleUtility) {
		super(logger);
		this.logger = logger;
		this.styleUtility = styleUtility;
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetDocumentReader createReader(
			final Stateful<SpreadsheetDocument> sfDocument)
					throws SpreadsheetException {
		return new OdsSimpleodfDocumentReader(this.styleUtility,
				new OdsSimpleodfStatefulDocument(sfDocument));
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetDocumentWriter createWriter(
			final Stateful<SpreadsheetDocument> sfDocument,
			final/*@Nullable*/OutputStream outputStream)
					throws SpreadsheetException {
		return new OdsSimpleodfDocumentWriter(this.logger, this.styleUtility,
				new OdsSimpleodfStatefulDocument(sfDocument), outputStream);
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetDocument loadSpreadsheetDocument(
			final InputStream inputStream) throws SpreadsheetException {
		SpreadsheetDocument document;
		try {
			document = SpreadsheetDocument.loadDocument(inputStream);
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		} catch (final Exception e) {
			throw new SpreadsheetException(e);
		}
		return document;
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetDocument newSpreadsheetDocument(
			final/*@Nullable*/OutputStream outputStream)
					throws SpreadsheetException {
		SpreadsheetDocument document;
		try {
			document = SpreadsheetDocument.newSpreadsheetDocument();
		} catch (final Exception e) {
			throw new SpreadsheetException(e);
		}
		return document;
	}
}
