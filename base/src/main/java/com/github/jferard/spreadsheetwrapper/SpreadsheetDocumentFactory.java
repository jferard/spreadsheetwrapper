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

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * A factory to create documents writers and readers
 */
public interface SpreadsheetDocumentFactory {

	/**
	 * Creates a workbook writer with no optionalOutput (use saveAs). May throw
	 * UnsupportedOperationException
	 *
	 * @return the writer
	 */
	SpreadsheetDocumentWriter create() throws SpreadsheetException;

	/**
	 * Creates a workbook writer
	 *
	 * @param outputFile
	 *            the file to write
	 * @return the writer
	 * */
	SpreadsheetDocumentWriter create(final File outputFile)
			throws SpreadsheetException;

	/**
	 * Creates a workbook writer
	 *
	 * @param outputStream
	 *            the stream to write to
	 * @return the writer
	 * */
	SpreadsheetDocumentWriter create(
			final/*@Nullable*/OutputStream outputStream)
			throws SpreadsheetException;

	/**
	 * creates a new File object
	 *
	 * @param parent
	 *            parent file
	 * @param childWithoutExtension
	 *            child filename without extension
	 * @return the File object
	 */
	File createNewFile(File parent, String childWithoutExtension);

	/**
	 * creates a new File object
	 *
	 * @param pathname
	 *            filename without extension
	 * @return the File object
	 */
	File createNewFile(String pathname);

	/**
	 * creates a new File object
	 *
	 * @param parent
	 *            parent directory name
	 * @param childWithoutExtension
	 *            child filename without extension
	 * @return the File object
	 */
	File createNewFile(String parent, String child);

	/**
	 * @return the extension for the files created and read by the factory
	 */
	String getExtension();

	/**
	 * Creates a workbook reader from an existing workbook
	 *
	 * @param inputFile
	 *            the file to read from
	 * @return the reader
	 * */
	SpreadsheetDocumentReader openForRead(final File inputFile)
			throws SpreadsheetException;

	/**
	 * Creates a workbook reader from an existing workbook
	 *
	 * @param inputStream
	 *            the stream to read from
	 * @return the reader
	 * */
	SpreadsheetDocumentReader openForRead(final InputStream inputStream)
			throws SpreadsheetException;

	/**
	 * Open a workbook writer from a existing workbook with no optionalOutput (use
	 * saveAs). May throw UnsupportedOperationException
	 *
	 * @param inputFile
	 *            the file to read from
	 * @return the writer
	 */
	SpreadsheetDocumentWriter openForWrite(final File inputFile)
			throws SpreadsheetException;

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
	SpreadsheetDocumentWriter openForWrite(final File inputFile,
			final File outputFile) throws SpreadsheetException;

	/**
	 * Open a workbook writer from a existing workbook with no optionalOutput (use
	 * saveAs). May throw UnsupportedOperationException
	 *
	 * @param inputStream
	 *            the stream to read from
	 * @return the writer
	 */
	SpreadsheetDocumentWriter openForWrite(final InputStream inputStream)
			throws SpreadsheetException;

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
			final/*@Nullable*/OutputStream outputStream)
			throws SpreadsheetException;
}