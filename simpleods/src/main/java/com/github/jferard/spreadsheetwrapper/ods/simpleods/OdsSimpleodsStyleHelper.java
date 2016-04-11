package com.github.jferard.spreadsheetwrapper.ods.simpleods;

import org.simpleods.BorderStyle;
import org.simpleods.OdsFile;
import org.simpleods.TableCell;
import org.simpleods.TableStyle;
import org.simpleods.TextStyle;

import com.github.jferard.spreadsheetwrapper.Util;
import com.github.jferard.spreadsheetwrapper.style.Borders;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;

public class OdsSimpleodsStyleHelper {

	public String getStyleName(TableCell simpleCell) {
		if (simpleCell == null)
			return null;

		return simpleCell.getStyle();
	}

	public boolean setStyle(OdsFile file, String styleName,
			WrapperCellStyle wrapperCellStyle) {
		final TableStyle newStyle = new TableStyle(TableStyle.STYLE_TABLECELL,
				styleName, file);
		final WrapperFont wrapperFont = wrapperCellStyle.getCellFont();
		if (wrapperFont != null) {
			final TextStyle textStyle = newStyle.getTextStyle();
			final WrapperColor fontColor = wrapperFont.getColor();
			final int bold = wrapperFont.getBold();
			final int italic = wrapperFont.getItalic();
			final double size = wrapperFont.getSize();
			final String family = wrapperFont.getFamily();

			if (fontColor != null)
				textStyle.setFontColor(fontColor.toHex());
			if (bold == WrapperCellStyle.YES) {
				textStyle.setFontWeightBold();
				if (italic == WrapperCellStyle.YES)
					textStyle.setFontWeightItalic();
			} else {
				if (italic == WrapperCellStyle.YES)
					textStyle.setFontWeightItalic();
				else
					textStyle.setFontWeightNormal();
			}
			if (!Util.almostEqual(size, WrapperCellStyle.DEFAULT))
				textStyle.setFontSize((int) size);
			if (family != null)
				textStyle.setFontName(family);

		}
		final WrapperColor backgroundColor = wrapperCellStyle
				.getBackgroundColor();
		final Borders borders = wrapperCellStyle.getBorders();
		if (backgroundColor != null)
			newStyle.setBackgroundColor(backgroundColor.toHex());
		if (borders != null) {
			final double borderLineWidth = borders.getLineWidth();
			if (borderLineWidth != WrapperCellStyle.DEFAULT)
				newStyle.addBorderStyle(Double.valueOf(borderLineWidth) + "pt",
						"#000000", BorderStyle.BORDER_SOLID,
						BorderStyle.POSITION_ALL);
		}
		return true;
	}

	public boolean setStyle(TableCell simpleCell, String styleName) {
		simpleCell.setStyle(styleName);
		return true;
	}

}
