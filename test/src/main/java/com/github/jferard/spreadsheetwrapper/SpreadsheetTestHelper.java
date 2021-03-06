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

public final class SpreadsheetTestHelper {
	/** the temp dir for writing files saved after each test */
	private static final String TEMP_DIR = System.getProperty("java.io.tmpdir");

	/**
	 * @return the caller class name
	 */
	public static String getCallerClassName() {
		final StackTraceElement[] stElements = Thread.currentThread()
				.getStackTrace();
		final StackTraceElement stElement = stElements[3];
		return stElement.getClassName();
	}

	/**
	 * @param factory
	 *            the factory
	 * @param className
	 *            the class name
	 * @param methodName
	 *            the method name
	 * @return a file name to identify the test
	 */
	public static File getOutputFile(final SpreadsheetDocumentFactory factory,
			final String className, final String methodName) {
		return factory.createNewFile(SpreadsheetTestHelper.TEMP_DIR,
				String.format("test-%s-%s", className, methodName));
	}

	private SpreadsheetTestHelper() {
	}
}
