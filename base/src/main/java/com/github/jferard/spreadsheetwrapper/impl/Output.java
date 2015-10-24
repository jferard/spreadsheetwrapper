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
	/** stream output */
	final/*@Nullable*/OutputStream outputStream;

	/** file output */
	final/*@Nullable*/File outputFile;

	/** url output */
	final/*@Nullable*/URL outputURL;

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

	public Output(OutputStream outputStream) {
		this(outputStream, null, null);
	}

	public Output(File outputFile) {
		this(null, outputFile, null);
	}

	public Output(URL outputURL) {
		this(null, null, outputURL);
	}

	public Output() {
		this(null, null, null);
	}

	private Output(OutputStream outputStream, File outputFile, URL outputURL) {
		super();
		this.outputStream = outputStream;
		this.outputFile = outputFile;
		this.outputURL = outputURL;
	}

	public/*@Nullable*/OutputStream getStream() {
		OutputStream outputStream;
		try {
			if (this.outputStream != null)
				outputStream = this.outputStream;
			else if (this.outputFile != null)
				outputStream = getOutputStream(this.outputFile);
			else if (this.outputURL != null)
				outputStream = getOutputStream(this.outputURL);
			else
				outputStream = null;
		} catch (FileNotFoundException e) {
			outputStream = null;
		} catch (IOException e) {
			outputStream = null;
		}
		return outputStream;
	}

}
