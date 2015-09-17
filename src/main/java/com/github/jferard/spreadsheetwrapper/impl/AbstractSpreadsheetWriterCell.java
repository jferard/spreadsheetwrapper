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

import java.util.Date;

import com.github.jferard.spreadsheetwrapper.SpreadsheetWriterCursor;

public abstract class AbstractSpreadsheetWriterCell implements
SpreadsheetWriterCursor {
	AbstractSpreadsheetWriterCell() {
		// does nothing
	}

	/** {@inheritDoc  */
	@Override
	public void setDate(final Date date, final String styleName) {
		this.setDate(date);
		this.setStyleName(styleName);
	}

	/** {@inheritDoc  */
	@Override
	public void setDouble(final Double value, final String styleName) {
		this.setDouble(value);
		this.setStyleName(styleName);
	}

	/** {@inheritDoc  */
	@Override
	public void setFormula(final String formula, final String styleName) {
		this.setFormula(formula);
		this.setStyleName(styleName);
	}

	/** {@inheritDoc  */
	@Override
	public void setInteger(final Integer value, final String styleName) {
		this.setInteger(value);
		this.setStyleName(styleName);
	}

	/** {@inheritDoc  */
	@Override
	public void setText(final String text, final String styleName) {
		this.setText(text);
		this.setStyleName(styleName);
	}

}
