package com.github.jferard.spreadsheetwrapper;

import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;

public interface StyleHelper<S, E> {
	WrapperCellStyle getWrapperCellStyle(E element);
	
	WrapperCellStyle toWrapperCellStyle(S internalStyle);
	
	void setWrapperCellStyle(E element, WrapperCellStyle wrapperCellStyle);
}
