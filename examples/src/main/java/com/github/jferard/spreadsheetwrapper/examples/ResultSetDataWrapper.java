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

import java.io.IOException;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.github.jferard.spreadsheetwrapper.DataWrapper;
import com.github.jferard.spreadsheetwrapper.SpreadsheetWriter;

public class ResultSetDataWrapper implements DataWrapper {
	private int columnCount;
	private final Logger logger;
	private final int max;
	private ResultSetMetaData metadata;
	private final ResultSet resultSet;

	public ResultSetDataWrapper(final Logger logger, final ResultSet rs,
			final int max) {
		this.logger = logger;
		this.resultSet = rs;
		this.max = max;
		try {
			this.metadata = rs.getMetaData();
			this.columnCount = this.metadata.getColumnCount();
		} catch (final SQLException e) {
			this.logger.log(Level.WARNING, "", e);
		}
	}

	/* {@inheritDoc} */
	@Override
	public boolean writeDataTo(final SpreadsheetWriter writer, final int r,
			final int c) {
		if (this.metadata == null)
			return false;

		int count = 0;
		try {
			if (this.resultSet.next()) {
				this.writeFirstLineDataTo(writer, r, c);
				int i = r;
				do {
					i++;
					count++;
					if (count <= this.max)
						this.writeDateLineTo(writer, i, c);
				} while (this.resultSet.next());
				

				this.writeLastLineDataTo(writer, r, c, count);
			}
		} catch (final SQLException e) {
			this.logger.log(Level.WARNING, "", e);
		} catch (final IOException e) {
			this.logger.log(Level.WARNING, "", e);
		}
		return count > 0;
	}

	/**
	 * @return the name of the columns
	 * @throws SQLException
	 */
	private List<String> getColumnNames() throws SQLException {
		final List<String> names = new ArrayList<String>(this.columnCount);
		for (int i = 0; i < this.columnCount; i++)
			names.add(this.metadata.getColumnName(i + 1));

		return names;
	}

	/**
	 * @return the values of the current column
	 * @throws SQLException
	 * @throws IOException
	 */
	private List<Object> getColumnValues() throws SQLException, IOException {
		final List<Object> values = new ArrayList<Object>(this.columnCount);
		for (int i = 0; i < this.metadata.getColumnCount(); i++) {
			values.add(this.resultSet.getObject(i + 1));
		}
		return values;
	}

	private void writeDateLineTo(final SpreadsheetWriter writer, final int i,
			final int c) throws SQLException, IOException {

		final List<Object> columnValues = this.getColumnValues();
		for (int j = c; j < c + this.columnCount; j++) {
			final Object object = columnValues.get(j);
			if (object == null)
				writer.setCellContent(i, j, "");
			else
				try {
					writer.setCellContent(i, j, object);
				} catch (final IllegalArgumentException e) {
					this.logger.fine(String.format("%s -> %s (%s)", e, object,
							object.getClass()));
					writer.setCellContent(i, j, object.toString());
				}
		}
	}

	private void writeFirstLineDataTo(final SpreadsheetWriter writer,
			final int r, final int c) throws SQLException {
		final List<String> columnNames = this.getColumnNames();
		for (int j = c; j < c + this.columnCount; j++) {
			writer.setCellContent(r, j, columnNames.get(j), "head");
		}
		this.logger.fine(String.format("Head %s", columnNames));
	}

	private void writeLastLineDataTo(final SpreadsheetWriter writer,
			final int r, final int c, final int count) {
		if (count == 0) {// no row
			for (int j = c; j < c + this.columnCount; j++)
				writer.setCellContent(r + count + 1, j, "");
			this.logger.fine("No row written");

		} else if (count > this.max) {
			for (int j = c; j < c + this.columnCount; j++) {
				writer.setCellContent(r + this.max + 1, j,
						String.format("... (%d rows remaining)", count - this.max));
			}
			this.logger.fine(String.format("%d total rows, %d written rows", count,
					count - this.max));
		} else
			this.logger.fine(String.format("%d written rows", count));
	}
}
