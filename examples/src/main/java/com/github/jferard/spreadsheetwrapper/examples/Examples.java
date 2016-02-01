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
package com.github.jferard.spreadsheetwrapper.examples;

import java.util.Arrays;
import java.util.logging.Logger;

import com.github.jferard.spreadsheetwrapper.DataWrapper;
import com.github.jferard.spreadsheetwrapper.DocumentFactoryManager;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentFactory;
import com.github.jferard.spreadsheetwrapper.SpreadsheetDocumentWriter;
import com.github.jferard.spreadsheetwrapper.SpreadsheetException;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;

public final class Examples {
	/**
	 * A DataWrapper example : one object per row : toString, class name, hash code
	 */
	private static final class DataWrapperExample implements DataWrapper {
		/** the object to be written */
		private final Object[] objects;

		/**
		 * @param objects the object to be written
		 */
		public DataWrapperExample(final Object... objects) {
			this.objects = objects;
		}

		/** {@inheritDoc} */
		@Override
		public boolean writeDataTo(final SpreadsheetWriter writer, final int r,
				final int c) {
			for (int i = 0; i < this.objects.length; i++) {
				final Object object = this.objects[i];
				writer.setText(r + i, c, object.toString());
				writer.setText(r + i, c + 1, object.getClass()
						.getName());
				writer.setInteger(r + i, c + 2, object.hashCode());
			}
			return true;
		}
	}

	/**
	 * First example. Needs odfdom in path.
	 *
	 * @throws SpreadsheetException
	 */
	public static void createSimpleDocument() throws SpreadsheetException {
		final DocumentFactoryManager manager = new DocumentFactoryManager(
				Logger.getAnonymousLogger());
		final SpreadsheetDocumentFactory factory = manager
				.getFactory("ods.odfdom.OdsOdfdomDocumentFactory");
		final SpreadsheetDocumentWriter documentWriter = factory.create();
		final SpreadsheetWriter newSheet = documentWriter
				.addSheet("first_sheet");
		newSheet.setText(0, 0, "First text");
		documentWriter.saveAs(factory.createNewFile(System.getProperty("java.io.tmpdir"), "createsimpledocument"));
	}

	/**
	 * A simple example :of the wrapper. @see DataWrapperExample class.
	 * @throws SpreadsheetException
	 */
	public static void wrapperExample() throws SpreadsheetException {
		final DataWrapper wrapper = new DataWrapperExample(
				Integer.valueOf(1), Double.valueOf(2.0), "3", Arrays.asList(4, 5, 6));

		final DocumentFactoryManager manager = new DocumentFactoryManager(
				Logger.getAnonymousLogger());
		final SpreadsheetDocumentFactory factory = manager
				.getFactory("ods.odfdom.OdsOdfdomDocumentFactory");
		final SpreadsheetDocumentWriter documentWriter = factory.create();
		final SpreadsheetWriter newSheet = documentWriter
				.addSheet("first_sheet");
		newSheet.setText(0, 0, "toString()");
		newSheet.setText(0, 1, "getClass().getName()");
		newSheet.setText(0, 2, "hashCode()");
		newSheet.writeDataFrom(1, 0, wrapper);
		documentWriter.saveAs(factory.createNewFile(System.getProperty("java.io.tmpdir"), "wrapperexample"));
	}

	private Examples() {
	}
}
