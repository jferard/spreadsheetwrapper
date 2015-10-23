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
package com.github.jferard.spreadsheetwrapper.xls.poi;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.impl.AbstractDocumentFactory;
import com.github.jferard.spreadsheetwrapper.impl.Stateful;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class XlsPoiDocumentFactory extends AbstractDocumentFactory<Workbook>
		implements SpreadsheetDocumentFactory {
	public static SpreadsheetDocumentFactory create(final Logger logger) {
		return new XlsPoiDocumentFactory(logger, new XlsPoiStyleUtility());
	}

	private final Logger logger;

	private final XlsPoiStyleUtility styleUtility;

	public XlsPoiDocumentFactory(final Logger logger,
			final XlsPoiStyleUtility styleUtility) {
		super(logger);
		this.logger = logger;
		this.styleUtility = styleUtility;
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetDocumentReader createReader(
			final Stateful<Workbook> sfWorkbook) throws SpreadsheetException {
		return new XlsPoiDocumentReader(this.styleUtility,
				sfWorkbook.getObject());
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetDocumentWriter createWriter(
			final Stateful<Workbook> sfWorkbook,
			final/*@Nullable*/OutputStream outputStream)
					throws SpreadsheetException {
		return new XlsPoiDocumentWriter(this.logger, this.styleUtility,
				sfWorkbook.getObject(), outputStream);
	}

	/** {@inheritDoc} */
	@Override
	protected Workbook loadSpreadsheetDocument(final InputStream inputStream)
			throws SpreadsheetException {
		Workbook workbook;
		try {
			workbook = WorkbookFactory.create(inputStream);
		} catch (final InvalidFormatException e) {
			throw new SpreadsheetException(e);
		} catch (final IOException e) {
			throw new SpreadsheetException(e);
		}
		return workbook;
	}

	/** {@inheritDoc} */
	@Override
	protected Workbook newSpreadsheetDocument(
	/*@Nullable*/final OutputStream outputStream) throws SpreadsheetException {
		final Workbook workbook = new HSSFWorkbook();
		return workbook;
	}
}
