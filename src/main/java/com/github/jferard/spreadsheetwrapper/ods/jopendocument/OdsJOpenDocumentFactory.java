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

package com.github.jferard.spreadsheetwrapper.ods.jopendocument;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.logging.Logger;

import javax.swing.table.DefaultTableModel;

import org.jopendocument.dom.ODPackage;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.impl.AbstractDocumentFactory;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class OdsJOpenDocumentFactory extends
		AbstractDocumentFactory<SpreadSheet> implements
		SpreadsheetDocumentFactory {
	/** simple logger */
	private final Logger logger;

	/**
	 * @param logger
	 *            simple logger
	 */
	public OdsJOpenDocumentFactory(final Logger logger) {
		super(logger);
		this.logger = logger;
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetDocumentReader createReader(final SpreadSheet document)
			throws SpreadsheetException {
		return new OdsJOpenDocumentReader(document);
	}

	/**
	 * {@inheritDoc}
	 *
	 * @throws SpreadsheetException
	 */
	@Override
	protected SpreadsheetDocumentWriter createWriter(
			final SpreadSheet document,
			final/*@Nullable*/OutputStream outputStream)
			throws SpreadsheetException {
		return new OdsJOpenDocumentWriter(this.logger, document, outputStream);
	}

	@Override
	protected SpreadSheet loadSpreadsheetDocument(final InputStream inputStream)
			throws SpreadsheetException {
		try {
			return new ODPackage(inputStream).getSpreadSheet();
		} catch (IOException e) {
			throw new SpreadsheetException(e);
		}
	}

	@Override
	protected SpreadSheet newSpreadsheetDocument(
			final/*@Nullable*/OutputStream outputStream)
			throws SpreadsheetException {
		try {
			return SpreadSheet.createEmpty(new DefaultTableModel());
		} catch (IOException e) {
			throw new SpreadsheetException(e);
		}
	}
}
