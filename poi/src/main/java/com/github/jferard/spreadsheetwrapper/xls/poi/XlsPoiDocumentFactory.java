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
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.impl.AbstractDocumentFactory;
import com.github.jferard.spreadsheetwrapper.impl.OptionalOutput;
import com.github.jferard.spreadsheetwrapper.style.CellStyleAccessor;
import com.github.jferard.spreadsheetwrapper.xls.XlsConstants;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class XlsPoiDocumentFactory extends AbstractDocumentFactory<Workbook>
		implements SpreadsheetDocumentFactory {
	/**
	 * static function for the manager
	 *
	 * @param logger
	 *            the logger
	 * @return a factory
	 */
	public static SpreadsheetDocumentFactory create(final Logger logger) {
		XlsPoiStyleColorHelper colorHelper = new XlsPoiStyleColorHelper(logger);
		return new XlsPoiDocumentFactory(logger, new XlsPoiStyleHelper(
				new CellStyleAccessor<CellStyle>(), colorHelper,
				new XlsPoiStyleFontHelper(colorHelper),
				new XlsPoiStyleBorderHelper(colorHelper)));
	}

	/** the logger */
	private final Logger logger;

	/** the style helper */
	private final XlsPoiStyleHelper styleHelper;

	/**
	 * @param logger
	 *            the logger
	 * @param styleHelper
	 *            the style helper
	 */
	public XlsPoiDocumentFactory(final Logger logger,
			final XlsPoiStyleHelper styleHelper) {
		super(XlsConstants.EXTENSION1);
		this.logger = logger;
		this.styleHelper = styleHelper;
	}

	/** {@inheritDoc} */
	@Override
	protected SpreadsheetDocumentWriter createWriter(
			final Workbook workbook, final OptionalOutput optionalOutput)
			throws SpreadsheetException {
		return new XlsPoiDocumentWriter(this.logger, workbook,
				this.styleHelper, optionalOutput);
	}

	/** {@inheritDoc} */
	@Override
	protected Workbook loadSpreadsheetDocument(final InputStream inputStream)
			throws SpreadsheetException {
		Workbook workbook;
		try {
			workbook = WorkbookFactory.create(inputStream);
			inputStream.close();
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
		return new HSSFWorkbook();
	}
}
