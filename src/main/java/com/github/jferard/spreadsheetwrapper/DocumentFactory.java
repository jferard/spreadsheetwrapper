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

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
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
	 * Creates a workbook writer
	 *
	 * @param outputFile
	 *            the file to write
	 * @return the writer
	 * */
	public SpreadsheetDocumentWriter create(final File outputFile)
			throws SpreadsheetException {
		return this.create(outputFile, this.getExtension(outputFile.getPath()));
	}

	/**
	 * Creates a workbook writer
	 *
	 * @param outputFile
	 *            the file to write
	 * @return the writer
	 * */
	public SpreadsheetDocumentWriter create(final File outputFile,
			final String extension) throws SpreadsheetException {
		final SpreadsheetDocumentFactory documentFactory = this
				.getDocumentFactory(extension);
		return documentFactory.create(outputFile);
	}

	/**
	 * Creates a workbook writer
	 *
	 * @param outputStream
	 *            the stream to write to
	 * @return the writer
	 * */
	public SpreadsheetDocumentWriter create(
			final/*@Nullable*/OutputStream outputStream, final String extension)
			throws SpreadsheetException {
		final SpreadsheetDocumentFactory documentFactory = this
				.getDocumentFactory(extension);
		return documentFactory.create(outputStream);
	}

	/**
	 * Creates a workbook writer with no output (use saveAs). May throw
	 * UnsupportedOperationException
	 */
	public SpreadsheetDocumentWriter create(final String extension)
			throws SpreadsheetException {
		final SpreadsheetDocumentFactory documentFactory = this
				.getDocumentFactory(extension);
		return documentFactory.create();
	}

	/**
	 * @param outputURL
	 * @return
	 * @throws SpreadsheetException
	 * @deprecated
	 */
	@Deprecated
	public SpreadsheetDocumentWriter create(final URL outputURL)
			throws SpreadsheetException {
		return this.create(outputURL, this.getExtension(outputURL.getPath()));
	}

	/**
	 * @param outputURL
	 * @return
	 * @throws SpreadsheetException
	 * @deprecated
	 */
	@Deprecated
	public SpreadsheetDocumentWriter create(final URL outputURL,
			final String extension) throws SpreadsheetException {
		final SpreadsheetDocumentFactory documentFactory = this
				.getDocumentFactory(extension);
		return documentFactory.create(outputURL);
	}

	/**
	 * Creates a workbook reader from an existing workbook
	 *
	 * @param inputFile
	 *            the file to read from
	 * @return the reader
	 * */
	public SpreadsheetDocumentReader openForRead(final File inputFile)
			throws SpreadsheetException {
		return this.openForRead(inputFile,
				this.getExtension(inputFile.getPath()));
	}

	/**
	 * Creates a workbook reader from an existing workbook
	 *
	 * @param inputFile
	 *            the file to read from
	 * @return the reader
	 * */
	public SpreadsheetDocumentReader openForRead(final File inputFile,
			final String extension) throws SpreadsheetException {
		final SpreadsheetDocumentFactory documentFactory = this
				.getDocumentFactory(extension);
		return documentFactory.openForRead(inputFile);
	}

	/**
	 * Creates a workbook reader from an existing workbook
	 *
	 * @param inputStream
	 *            the stream to read from
	 * @return the reader
	 * */
	public SpreadsheetDocumentReader openForRead(final InputStream inputStream,
			final String extension) throws SpreadsheetException {
		final SpreadsheetDocumentFactory documentFactory = this
				.getDocumentFactory(extension);
		return documentFactory.openForRead(inputStream);
	}

	/**
	 * @param inputURL
	 *            URL to read from
	 * @return a reader on the document
	 * @throws SpreadsheetException
	 * @deprecated
	 */
	@Deprecated
	public SpreadsheetDocumentReader openForRead(final URL inputURL)
			throws SpreadsheetException {
		return this
				.openForRead(inputURL, this.getExtension(inputURL.getPath()));
	}

	/**
	 * @param inputURL
	 *            URL to read from
	 * @return a reader on the document
	 * @throws SpreadsheetException
	 * @deprecated
	 */
	@Deprecated
	public SpreadsheetDocumentReader openForRead(final URL inputURL,
			final String extension) throws SpreadsheetException {
		final SpreadsheetDocumentFactory documentFactory = this
				.getDocumentFactory(extension);
		return documentFactory.openForRead(inputURL);
	}

	/**
	 * Open a workbook writer from a existing workbook with no output (use
	 * saveAs). May throw UnsupportedOperationException
	 *
	 * @param inputFile
	 *            the file to read from
	 * @return the writer
	 */
	public SpreadsheetDocumentWriter openForWrite(final File inputFile)
			throws SpreadsheetException {
		return this.openForWrite(inputFile,
				this.getExtension(inputFile.getPath()));
	}

	/**
	 * Creates a workbook writer from an existing workbook
	 *
	 * @param inputFile
	 *            the file to read from
	 * @param outputFile
	 *            the file to write to
	 * @return the writer
	 * @throws SpreadsheetException
	 *             if can't open the writer
	 */
	public SpreadsheetDocumentWriter openForWrite(final File inputFile,
			final File outputFile) throws SpreadsheetException {
		return this.openForWrite(inputFile, outputFile,
				this.getExtension(inputFile.getPath()));
	}

	/**
	 * Creates a workbook writer from an existing workbook
	 *
	 * @param inputFile
	 *            the file to read from
	 * @param outputFile
	 *            the file to write to
	 * @return the writer
	 * @throws SpreadsheetException
	 *             if can't open the writer
	 */
	public SpreadsheetDocumentWriter openForWrite(final File inputFile,
			final File outputFile, final String extension)
			throws SpreadsheetException {
		final SpreadsheetDocumentFactory documentFactory = this
				.getDocumentFactory(extension);
		return documentFactory.openForWrite(inputFile, outputFile);
	}

	/**
	 * Open a workbook writer from a existing workbook with no output (use
	 * saveAs). May throw UnsupportedOperationException
	 *
	 * @param inputFile
	 *            the file to read from
	 * @return the writer
	 */
	public SpreadsheetDocumentWriter openForWrite(final File inputFile,
			final String extension) throws SpreadsheetException {
		final SpreadsheetDocumentFactory documentFactory = this
				.getDocumentFactory(extension);
		return documentFactory.openForWrite(inputFile);
	}

	/**
	 * Open a workbook writer from a existing workbook with no output (use
	 * saveAs). May throw UnsupportedOperationException
	 *
	 * @param inputStream
	 *            the stream to read from
	 * @return the writer
	 */
	public SpreadsheetDocumentWriter openForWrite(
			final InputStream inputStream, final String extension)
			throws SpreadsheetException {
		final SpreadsheetDocumentFactory documentFactory = this
				.getDocumentFactory(extension);
		return documentFactory.openForWrite(inputStream);

	}

	/**
	 * @param inputURL
	 *            URL to read from
	 * @return a writer on the document
	 * @throws SpreadsheetException
	 * @deprecated
	 */
	@Deprecated
	public SpreadsheetDocumentWriter openForWrite(final URL inputURL)
			throws SpreadsheetException {
		return this.openForWrite(inputURL,
				this.getExtension(inputURL.getPath()));
	}

	/**
	 * @param inputURL
	 *            URL to read from
	 * @return a writer on the document
	 * @throws SpreadsheetException
	 * @deprecated
	 */
	@Deprecated
	public SpreadsheetDocumentWriter openForWrite(final URL inputURL,
			final String extension) throws SpreadsheetException {
		final SpreadsheetDocumentFactory documentFactory = this
				.getDocumentFactory(extension);
		return documentFactory.openForWrite(inputURL);
	}

	/**
	 * @param inputURL
	 *            URL to read from
	 * @param outputURL
	 *            URL to write to
	 * @return a writer on the document
	 * @throws SpreadsheetException
	 * @deprecated
	 */
	@Deprecated
	public SpreadsheetDocumentWriter openForWrite(final URL inputURL,
			final URL outputURL) throws SpreadsheetException {
		return this.openForWrite(inputURL, outputURL,
				this.getExtension(inputURL.getPath()));
	}

	/**
	 * @param inputURL
	 *            URL to read from
	 * @param outputURL
	 *            URL to write to
	 * @return a writer on the document
	 * @throws SpreadsheetException
	 * @deprecated
	 */
	@Deprecated
	public SpreadsheetDocumentWriter openForWrite(final URL inputURL,
			final URL outputURL, final String extension)
			throws SpreadsheetException {
		final SpreadsheetDocumentFactory documentFactory = this
				.getDocumentFactory(extension);
		return documentFactory.openForWrite(inputURL, outputURL);
	}

	/**
	 * Creates a workbook writer
	 *
	 * @param outputFile
	 *            the file to write
	 * @return the writer
	 * */
	private SpreadsheetDocumentFactory getDocumentFactory(final String extension)
			throws SpreadsheetException {
		final SpreadsheetDocumentFactory documentFactory;
		if (extension.equals("ods")) {
			documentFactory = this.odsDocumentFactory;
		} else if (extension.equals("xls") || extension.equals("xlsx"))
			documentFactory = this.xlsDocumentFactory;
		else
			throw new SpreadsheetException(
					"Can't find extension, use createOds or createXls");

		return documentFactory;
	}

	private String getExtension(final String filePath)
			throws SpreadsheetException {
		final int dotIndex = filePath.lastIndexOf('.');
		if (dotIndex == -1)
			throw new SpreadsheetException(
					"Can't find extension, use ...Ods(...) or ...Xls(...)");

		final String extension = filePath.substring(dotIndex + 1);
		return extension;
	}

	/**
	 * Creates a workbook writer from an existing workbook
	 *
	 * @param inputStream
	 *            the stream to read from
	 * @param outputStream
	 *            the stream to write to
	 * @return the writer
	 * @throws SpreadsheetException
	 *             if can't open the writer
	 */
	SpreadsheetDocumentWriter openForWrite(final InputStream inputStream,
			final/*@Nullable*/OutputStream outputStream, final String extension)
					throws SpreadsheetException {
		final SpreadsheetDocumentFactory documentFactory = this
				.getDocumentFactory(extension);
		return documentFactory.openForWrite(inputStream, outputStream);
	}

}
