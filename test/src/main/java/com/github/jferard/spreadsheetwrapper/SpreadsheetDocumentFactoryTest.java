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
public abstract class SpreadsheetDocumentFactoryTest {
	/** Name of the test */
	@Rule
	public TestName name = new TestName();

	/** the destination file (test are written in a temp directory) */
	private File destFile;

	/** the source file (from test/resources) */
	private File sourceFile;

	/** the source URL (from test/resources) */
	private URL sourceURL;

	/** the factory */
	protected SpreadsheetDocumentFactory factory;

	public SpreadsheetDocumentFactoryTest() {
		super();
	}

	@Before
	@SuppressWarnings("nullness")
	public void setUp() throws URISyntaxException {
		this.factory = this.getProperties().getFactory();
		this.sourceURL = this.getProperties().getSourceURL();
		Assume.assumeNotNull(this.sourceURL);

		this.sourceFile = new File(this.sourceURL.toURI());
		this.destFile = SpreadsheetTestHelper.getOutputFile(this.getClass()
				.getSimpleName(), this.name.getMethodName(), this
				.getProperties().getExtension());
	}

	@Test
	public void testCreateEmptyDocumentWithDestinationFile() {
		try {
			final SpreadsheetDocumentWriter sdw = this.factory
					.create(this.destFile);
			sdw.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public void testCreateEmptyDocumentWithNoDestination() {
		try {
			final SpreadsheetDocumentWriter sdw = this.factory.create();
			sdw.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public final void testOpenFileForRead() {
		try {
			final SpreadsheetDocumentReader sdr = this.factory
					.openForRead(this.sourceFile);
			sdr.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public void testOpenFileForWriteWithDest() {
		try {
			final SpreadsheetDocumentReader sdr = this.factory.openForWrite(
					this.sourceFile, this.destFile);
			sdr.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public void testOpenFileForWriteWithNoDest() {
		try {
			final SpreadsheetDocumentWriter sdw = this.factory
					.openForWrite(this.sourceFile);
			sdw.save();
			sdw.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public final void testOpenStreamForRead() {
		try {
			final InputStream sourceStream = this.sourceURL.openStream();
			final SpreadsheetDocumentReader sdr = this.factory
					.openForRead(sourceStream);
			sdr.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		} catch (final IOException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public void testOpenStreamForWriteWithDest() {
		try {
			final InputStream sourceStream = this.sourceURL.openStream();
			final OutputStream destStream = new FileOutputStream(this.destFile);
			final SpreadsheetDocumentWriter sdw = this.factory.openForWrite(
					sourceStream, destStream);
			sdw.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		} catch (final IOException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	@Test
	public void testOpenStreamForWriteWithNoDest() {
		try {
			final InputStream sourceStream = this.sourceURL.openStream();
			final SpreadsheetDocumentReader sdr = this.factory
					.openForWrite(sourceStream);
			sdr.close();
		} catch (final SpreadsheetException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		} catch (final IOException e) {
			e.printStackTrace();
			Assert.fail(e.getMessage());
		} catch (final UnsupportedOperationException e) {
			Assume.assumeNoException(e);
		}
	}

	protected abstract TestProperties getProperties();
}