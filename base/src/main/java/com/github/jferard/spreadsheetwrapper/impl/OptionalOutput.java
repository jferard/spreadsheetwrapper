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
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * The Ouput class stores one of the possible optionalOutput (stream, file,
 * URL). This is a way to delay the creation of the outputStream. (If the
 * outputStream is created immediately and the optionalOutput file is the same
 * as the input file, then the input file will be erased before the spreadsheet
 * is opened).
 *
 * @author Julien
 *
 */
public class OptionalOutput {
	/**
	 * Empty optionalOutput
	 */
	public static OptionalOutput EMPTY = new OptionalOutput(null, null);

	/**
	 * OptionalOutput to a file
	 *
	 * @param outputFile
	 */
	public static OptionalOutput fromFile(final/*@Nullable*/File outputFile) {
		return new OptionalOutput(null, outputFile);
	}

	/**
	 * OptionalOutput to a stream
	 *
	 * @param outputStream
	 */
	public static OptionalOutput fromStream(
			final/*@Nullable*/OutputStream outputStream) {
		return new OptionalOutput(outputStream, null);
	}

	/** file optionalOutput */
	private final/*@Nullable*/File outputFile;

	/** stream optionalOutput */
	private/*@Nullable*/OutputStream outputStream;

	/**
	 * OptionalOutput to file or stream
	 *
	 * @param outputStream
	 * @param outputFile
	 */
	private OptionalOutput(final/*@Nullable*/OutputStream outputStream,
			final/*@Nullable*/File outputFile) {
		this.outputStream = outputStream;
		this.outputFile = outputFile;
	}

	/**
	 * Close the stream if opened
	 *
	 * @throws IOException
	 *             if can't close the stream
	 */
	public void close() throws IOException {
		if (this.outputStream != null)
			this.outputStream.close();
	}

	/**
	 * Get or open the stream
	 *
	 * @return the stream or null if it can't be opened
	 */
	public/*@Nullable*/OutputStream getStream() {
		if (this.outputStream == null && this.outputFile != null) {
			try {
				this.outputStream = new FileOutputStream(this.outputFile);
			} catch (final FileNotFoundException e) {
			}
		}
		return this.outputStream;
	}
}
