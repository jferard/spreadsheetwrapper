package com.github.jferard.spreadsheetwrapper.xls.jxl;

import java.io.IOException;
import jxl.Sheet;
import jxl.write.WritableSheet;
import jxl.write.WriteException;

public interface JxlWorkbook {

	void close() throws WriteException, IOException;

	WritableSheet createSheet(String sheetName, int index);

	int getNumberOfSheets();

	Sheet getSheet(String sheetName);

	String[] getSheetNames();

	Sheet[] getSheets();

	void write() throws IOException;
}
