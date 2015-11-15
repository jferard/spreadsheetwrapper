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

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.logging.Logger;

public class DocumentFactoryManager {
	/** simple logger */
	private final Logger logger;

	/**
	 * @param logger
	 *            simple logger
	 */
	public DocumentFactoryManager(final Logger logger) {
		super();
		this.logger = logger;
	}

	/**
	 * get a factory
	 *
	 * @throws ClassNotFoundException
	 * @throws IllegalAccessException
	 * @throws InstantiationException
	 * @throws InvocationTargetException
	 * @throws IllegalArgumentException
	 * @throws SecurityException
	 * @throws NoSuchMethodException
	 * */
	@SuppressWarnings("nullness")
	public SpreadsheetDocumentFactory getFactory(final String endOfClassName)
			throws SpreadsheetException {
		final Package package1 = this.getClass().getPackage();
		if (package1 == null)
			throw new SpreadsheetException("Package not found :"
					+ this.getClass());
		final String startOfClassName = package1.getName();
		Class<?> clazz;
		try {
			clazz = Class.forName(new StringBuilder(startOfClassName)
					.append('.').append(endOfClassName).toString());
			final Method method = clazz.getMethod("create", Logger.class);

			return (SpreadsheetDocumentFactory) method
					.invoke(null, this.logger);
		} catch (final ClassNotFoundException e) {
			throw new SpreadsheetException(e);
		} catch (final NoSuchMethodException e) {
			throw new SpreadsheetException(e);
		} catch (final SecurityException e) {
			throw new SpreadsheetException(e);
		} catch (final IllegalAccessException e) {
			throw new SpreadsheetException(e);
		} catch (final IllegalArgumentException e) {
			throw new SpreadsheetException(e);
		} catch (final InvocationTargetException e) {
			throw new SpreadsheetException(e);
		}
	}

}
