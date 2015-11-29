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
package com.github.jferard.spreadsheetwrapper.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentReader;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public abstract class AbstractBasicDocumentFactory implements
		SpreadsheetDocumentFactory {

	/** the extension */
	private final String extension;

	/**
	 * @param extension
	 *            the extension (ods or xls)
	 */
	protected AbstractBasicDocumentFactory(final String extension) {
		this.extension = extension;
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create() throws SpreadsheetException {
		return this.create((OutputStream) null);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter create(final/*@Nullable*/File outputFile)
			throws SpreadsheetException {
		if (outputFile == null)
			return this.create();

		try {
			final FileOutputStream fileOutputStream = new FileOutputStream(
					outputFile);
			return this.create(fileOutputStream);
		} catch (final FileNotFoundException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public File createNewFile(final File parent, final String child) {
		return new File(parent, child + "." + this.getExtension());
	}

	/** {@inheritDoc} */
	@Override
	public File createNewFile(final String pathname) {
		return new File(pathname + "." + this.getExtension());
	}

	/** {@inheritDoc} */
	@Override
	public File createNewFile(final String parent, final String child) {
		return new File(parent, child + "." + this.getExtension());
	}

	/** {@inheritDoc} */
	@Override
	public String getExtension() {
		return this.extension;
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentReader openForRead(final File inputFile)
			throws SpreadsheetException {
		try {
			return this.openForRead(new FileInputStream(inputFile));
		} catch (final FileNotFoundException e) {
			throw new SpreadsheetException(e);
		}
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter openForWrite(final File inputFile)
			throws SpreadsheetException {
		return this.openForWrite(inputFile, inputFile);
	}

	/** {@inheritDoc} */
	@Override
	public SpreadsheetDocumentWriter openForWrite(final InputStream inputStream)
			throws SpreadsheetException {
		return this.openForWrite(inputStream, (OutputStream) null);
	}
}