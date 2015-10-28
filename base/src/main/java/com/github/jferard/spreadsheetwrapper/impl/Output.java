package com.github.jferard.spreadsheetwrapper.impl;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URL;
import java.net.URLConnection;
import java.net.UnknownServiceException;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * The Ouput class stores one of the possible output (stream, file, URL). This
 * is a way to delay the creation of the outputStream. (If the outputStream is
 * created immediately and the output file is the same as the input file, then
 * the input file while be erased before the spreadsheet is opened).
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

	/**
	 * @param outputURL
	 *            the URL to open for write
	 * @return the output stream on this URL
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public static OutputStream getOutputStream(final URL outputURL)
			throws IOException, FileNotFoundException {
		OutputStream outputStream;

		final URLConnection connection = outputURL.openConnection();
		connection.setDoOutput(true);
		try {
			outputStream = connection.getOutputStream(); // NOPMD by Julien on
			// 30/08/15 12:54
		} catch (final UnknownServiceException e) {
			outputStream = new FileOutputStream(outputURL.getPath());
		}
		return outputStream;
	}

	/** file output */
	final/*@Nullable*/File outputFile;

	/** stream output */
	/*@Nullable*/OutputStream outputStream;

	/** url output */
	final/*@Nullable*/URL outputURL;

	public Output() {
		this(null, null, null);
	}

	public Output(final /*@Nullable*/ File outputFile) {
		this(null, outputFile, null);
	}

	public Output(final /*@Nullable*/ OutputStream outputStream) {
		this(outputStream, null, null);
	}

	public Output(final /*@Nullable*/ URL outputURL) {
		this(null, null, outputURL);
	}

	private Output(final /*@Nullable*/ OutputStream outputStream, final /*@Nullable*/ File outputFile,
			final /*@Nullable*/ URL outputURL) {
		super();
		this.outputStream = outputStream;
		this.outputFile = outputFile;
		this.outputURL = outputURL;
	}

	public void close() throws IOException {
		if (this.outputStream != null)
			this.outputStream.close();
	}

	public /*@Nullable*/ OutputStream getStream() {
		if (this.outputStream == null) {
			try {
				if (this.outputFile != null)
					this.outputStream = Output.getOutputStream(this.outputFile);
				else if (this.outputURL != null)
					this.outputStream = Output.getOutputStream(this.outputURL);
			} catch (final FileNotFoundException e) {
				// do nothing
			} catch (final IOException e) {
				// do nothing
			}
		}
		return this.outputStream;
	}
}
