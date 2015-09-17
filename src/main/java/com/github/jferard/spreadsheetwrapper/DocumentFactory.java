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
package com.github.jferard.spreadsheetwrapper;

import java.net.URL;
import java.util.logging.Logger;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class DocumentFactory {
	/** simple logger */
	private final Logger logger;
	private final SpreadsheetDocumentFactory odsDocumentFactory;
	private final SpreadsheetDocumentFactory xlsDocumentFactory;

	/**
	 * @param logger
	 *            simple logger
	 */
	public DocumentFactory(final Logger logger,
			final SpreadsheetDocumentFactory odsDocumentFactory,
			final SpreadsheetDocumentFactory xlsDocumentFactory) {
		super();
		this.logger = logger;
		this.odsDocumentFactory = odsDocumentFactory;
		this.xlsDocumentFactory = xlsDocumentFactory;
	}

	/**
	 * @param outputURL
	 * @return
	 * @throws SpreadsheetException
	 */
	@Deprecated
	public SpreadsheetDocumentWriter create(final URL outputURL)
			throws SpreadsheetException {
		final String fileName = outputURL.getPath();
		final int dotIndex = fileName.lastIndexOf('.');
		if (dotIndex == -1)
			throw new SpreadsheetException(
					"Can't find extension, use createOds or createXls");

		final String extension = fileName.substring(dotIndex + 1);
		final SpreadsheetDocumentWriter spreadsheetWriter;

		if (extension.equals("ods")) {
			spreadsheetWriter = this.createOds(outputURL);
		} else if (extension.equals("xls") || extension.equals("xlsx"))
			spreadsheetWriter = this.createXls(outputURL);
		else
			throw new SpreadsheetException(
					"Can't find extension, use createOds or createXls");

		return spreadsheetWriter;
	}

	/**
	 * Creates a workbook writer with no output (use saveAs). May throw
	 * UnsupportedOperationException
	 */
	public SpreadsheetDocumentWriter createOds() throws SpreadsheetException {
		return this.odsDocumentFactory.create();
	}

	@Deprecated
	public SpreadsheetDocumentWriter createOds(final URL outputURL)
			throws SpreadsheetException {
		return this.odsDocumentFactory.create(outputURL);
	}

	/**
	 * Creates a workbook writer with no output (use saveAs). May throw
	 * UnsupportedOperationException
	 */
	public SpreadsheetDocumentWriter createXls() throws SpreadsheetException {
		return this.xlsDocumentFactory.create();
	}

	@Deprecated
	public SpreadsheetDocumentWriter createXls(final URL outputURL)
			throws SpreadsheetException {
		return this.xlsDocumentFactory.create(outputURL);
	}

	@Deprecated
	public SpreadsheetDocumentWriter openCopyOf(final URL inputURL)
			throws SpreadsheetException {
		final String fileName = inputURL.getPath();
		final int dotIndex = fileName.lastIndexOf('.');
		if (dotIndex == -1)
			throw new SpreadsheetException(
					"Can't find extension, use openCopyOfOds or openCopyOfXls");

		final String extension = fileName.substring(dotIndex + 1);
		final SpreadsheetDocumentWriter spreadsheetWriter;

		if (extension.equals("ods")) {
			spreadsheetWriter = this.openCopyOfOds(inputURL);
		} else if (extension.equals("xls") || extension.equals("xlsx"))
			spreadsheetWriter = this.openCopyOfXls(inputURL);
		else
			throw new SpreadsheetException(
					"Can't find extension, use openCopyOfOds or openCopyOfXls");

		return spreadsheetWriter;
	}

	@Deprecated
	public SpreadsheetDocumentWriter openCopyOf(final URL inputURL,
			final URL outputURL) throws SpreadsheetException {
		final String fileName = inputURL.getPath();
		final int dotIndex = fileName.lastIndexOf('.');
		if (dotIndex == -1)
			throw new SpreadsheetException(
					"Can't find extension, use openCopyOfOds or openCopyOfXls");

		final String extension = fileName.substring(dotIndex + 1);
		final SpreadsheetDocumentWriter spreadsheetWriter;

		if (extension.equals("ods")) {
			spreadsheetWriter = this.openCopyOfOds(inputURL, outputURL);
		} else if (extension.equals("xls") || extension.equals("xlsx"))
			spreadsheetWriter = this.openCopyOfXls(inputURL, outputURL);
		else
			throw new SpreadsheetException(
					"Can't find extension, use openCopyOfOds or openCopyOfXls");

		return spreadsheetWriter;
	}

	@Deprecated
	public SpreadsheetDocumentWriter openCopyOfOds(final URL inputURL)
			throws SpreadsheetException {
		return this.odsDocumentFactory.openForWrite(inputURL);
	}

	@Deprecated
	public SpreadsheetDocumentWriter openCopyOfOds(final URL inputURL,
			final URL outputURL) throws SpreadsheetException {
		return this.odsDocumentFactory.openForWrite(inputURL, outputURL);
	}

	@Deprecated
	public SpreadsheetDocumentWriter openCopyOfXls(final URL inputURL,
			final URL outputURL) throws SpreadsheetException {
		return this.xlsDocumentFactory.openForWrite(inputURL, outputURL);
	}

	@Deprecated
	public SpreadsheetDocumentReader openForRead(final URL inputURL)
			throws SpreadsheetException {
		final String fileName = inputURL.getPath();
		final int dotIndex = fileName.lastIndexOf('.');
		if (dotIndex == -1)
			throw new SpreadsheetException(
					"Can't find extension, use openForReadOds or openForReadXls");

		final String extension = fileName.substring(dotIndex + 1);
		final SpreadsheetDocumentReader spreadsheetReader;

		if (extension.equals("ods")) {
			spreadsheetReader = this.openForReadOds(inputURL);
		} else if (extension.equals("xls") || extension.equals("xlsx"))
			spreadsheetReader = this.openForReadXls(inputURL);
		else
			throw new SpreadsheetException(
					"Can't find extension, use openForReadOds or openForReadXls");

		return spreadsheetReader;
	}

	@Deprecated
	public SpreadsheetDocumentReader openForReadOds(final URL inputURL)
			throws SpreadsheetException {
		return this.odsDocumentFactory.openForRead(inputURL);
	}

	@Deprecated
	public SpreadsheetDocumentReader openForReadXls(final URL inputURL)
			throws SpreadsheetException {
		return this.xlsDocumentFactory.openForRead(inputURL);
	}

	@Deprecated
	private SpreadsheetDocumentWriter openCopyOfXls(final URL inputURL)
			throws SpreadsheetException {
		return this.xlsDocumentFactory.openForWrite(inputURL);
	}
}
