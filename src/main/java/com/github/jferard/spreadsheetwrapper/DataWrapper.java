package com.github.jferard.spreadsheetwrapper;

public interface DataWrapper {
	boolean writeDataTo(SpreadsheetWriter writer, int r, int c);
}
