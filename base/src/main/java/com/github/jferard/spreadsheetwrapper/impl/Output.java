package com.github.jferard.spreadsheetwrapper.impl;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * The Ouput class stores one of the possible output (stream, file, URL). This
 * is a way to delay the creation of the outputStream. (If the outputStream is
 * created immediately and the output file is the same as the input file, then
 * the input file while be erased before the spreadsheet is opened).
 */
/**
 * @author Julien
 *
 */
public class Output {
	/**
	 * @param outputFile
	 *            the file to open for write
	 * @return the output stream on this URL
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public static OutputStream getOutputStream(final File outputFile)
			throws FileNotFoundException {
		return new FileOutputStream(outputFile);
	}

	/** file output */
	private final/*@Nullable*/File outputFile;

	/** stream output */
	private/*@Nullable*/OutputStream outputStream;

	/**
	 * Empty output
	 */
	public Output() {
		this(null, null);
	}

	/**
	 * Output to a file
	 *
	 * @param outputFile
	 */
	public Output(final/*@Nullable*/File outputFile) {
		this(null, outputFile);
	}

	/**
	 * Output to a stream
	 *
	 * @param outputStream
	 */
	public Output(final/*@Nullable*/OutputStream outputStream) {
		this(outputStream, null);
	}

	/**
	 * Output to file or stream
	 *
	 * @param outputStream
	 * @param outputFile
	 */
	private Output(final/*@Nullable*/OutputStream outputStream,
			final/*@Nullable*/File outputFile) {
		super();
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
		if (this.outputStream == null) {
			try {
				if (this.outputFile != null)
					this.outputStream = Output.getOutputStream(this.outputFile);
			} catch (final FileNotFoundException e) { // NOPMD by Julien on
				// 21/11/15 11:22
				// do nothing
			}
		}
		return this.outputStream;
	}
}
