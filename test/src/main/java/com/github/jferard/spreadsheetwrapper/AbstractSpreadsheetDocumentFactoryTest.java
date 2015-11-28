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
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.junit.Assert;
import org.junit.Assume;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TestName;

/**
 * Okay for all wrappers
 *
 */
public abstract class AbstractSpreadsheetDocumentFactoryTest {
	/** Name of the test */
	@Rule
	public TestName name = new TestName();

	/** the destination file (test are written in a temp directory) */
	private File destFile;

	/** logger, static initialization */
	private final Logger logger = Logger.getLogger(this.getClass().getName());

	/** the source file (from test/resources) */
	private File sourceFile;

	/** the source URL (from test/resources) */
	private URL sourceURL;

	/** the factory */
	protected SpreadsheetDocumentFactory factory;

	/** set the test up */
	@Before
	@SuppressWarnings("nullness")
	public void setUp() throws URISyntaxException {
		this.factory = this.getProperties().getFactory();
		this.sourceURL = this.getProperties().getResourceURL();
		Assume.assumeNotNull(this.sourceURL);

		this.sourceFile = new File(this.sourceURL.toURI());
		this.destFile = SpreadsheetTestHelper.getOutputFile(this.factory, this
				.getClass().getSimpleName(), this.name.getMethodName());
	}

	/** destination as file */
	@Test
	public void testCreateEmptyDocumentWithDestinationFile() {
		try {
			final SpreadsheetDocumentWriter documentWriter = this.factory
					.create(this.destFile);
			documentWriter.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	/** no destination */
	@Test
	public void testCreateEmptyDocumentWithNoDestination() {
		try {
			final SpreadsheetDocumentWriter documentWriter = this.factory
					.create();
			documentWriter.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	/** open existing file for read */
	@Test
	public final void testOpenFileForRead() {
		try {
			final SpreadsheetDocumentReader documentReader = this.factory
					.openForRead(this.sourceFile);
			documentReader.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	/** open existing file for write and define destination */
	@Test
	public void testOpenFileForWriteWithDest() {
		try {
			final SpreadsheetDocumentReader documentReader = this.factory
					.openForWrite(this.sourceFile, this.destFile);
			documentReader.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	/** open existing file for write but don't define destination */
	@Test
	public void testOpenFileForWriteWithNoDest() {
		try {
			final SpreadsheetDocumentWriter documentWriter = this.factory
					.openForWrite(this.sourceFile);
			documentWriter.save();
			documentWriter.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	/** open existing file for read, from stream */
	@Test
	public final void testOpenStreamForRead() {
		try {
			final InputStream sourceStream = this.sourceURL.openStream();
			final SpreadsheetDocumentReader documentReader = this.factory
					.openForRead(sourceStream);
			documentReader.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		} catch (final IOException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	/** open existing file for write ith destination stream */
	@Test
	public void testOpenStreamForWriteWithDest() {
		try {
			final InputStream sourceStream = this.sourceURL.openStream();
			final OutputStream destStream = new FileOutputStream(this.destFile);
			final SpreadsheetDocumentWriter documentWriter = this.factory
					.openForWrite(sourceStream, destStream);
			documentWriter.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		} catch (final IOException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	/** open existing file for write, but don't define dest, from stream */
	@Test
	public void testOpenStreamForWriteWithNoDest() {
		try {
			final InputStream sourceStream = this.sourceURL.openStream();
			final SpreadsheetDocumentReader documentReader = this.factory
					.openForWrite(sourceStream);
			documentReader.close();
		} catch (final SpreadsheetException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		} catch (final IOException e) {
			this.logger.log(Level.WARNING, "", e);
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	/** get the properties for tests */
	protected abstract TestProperties getProperties();
}