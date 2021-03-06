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
	public Date setDate(final Date date, final String styleName) {
		this.setStyleName(styleName);
		return this.setDate(date);
	}

	/** {@inheritDoc  */
	@Override
	public Double setDouble(final Number value, final String styleName) {
		this.setStyleName(styleName);
		return this.setDouble(value);
	}

	/** {@inheritDoc  */
	@Override
	public String setFormula(final String formula, final String styleName) {
		this.setStyleName(styleName);
		return this.setFormula(formula);
	}

	/** {@inheritDoc  */
	@Override
	public Integer setInteger(final Number value, final String styleName) {
		this.setStyleName(styleName);
		return this.setInteger(value);
	}

	/** {@inheritDoc  */
	@Override
	public String setText(final String text, final String styleName) {
		this.setStyleName(styleName);
		return this.setText(text);
	}

}
