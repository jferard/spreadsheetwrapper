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
import java.net.MalformedURLException;
import java.net.URL;

public class SpreadsheetTestHelper {
	static int i = 0;

	public static String getCallerClassName() {
		final StackTraceElement[] stElements = Thread.currentThread()
				.getStackTrace();
		final StackTraceElement stElement = stElements[3];
		return stElement.getClassName();
	}

	public static File getOutputFile(final String className,
			final String methodName, final String extension) {
		return new File(System.getProperty("java.io.tmpdir"), String.format(
				"test-%s-%s.%s", className, methodName, extension));
	}

	public static URL getOutputURL(final String extension)
			throws MalformedURLException {
		return new URL(String.format("file:///%s/%s", System
				.getProperty("java.io.tmpdir"), String.format("test-%s.%d.%s",
						SpreadsheetTestHelper.getCallerClassName(),
				SpreadsheetTestHelper.i++, extension)));
	}

	public static URL getOutputURL(final String className,
			final String methodName, final String extension)
					throws MalformedURLException {
		return new URL(String.format("file:///%s/%s", System
				.getProperty("java.io.tmpdir"), String.format("test-%s-%s.%s",
						className, methodName, extension)));
	}

}