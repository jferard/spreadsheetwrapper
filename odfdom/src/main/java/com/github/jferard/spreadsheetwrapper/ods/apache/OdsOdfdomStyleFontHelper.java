package com.github.jferard.spreadsheetwrapper.ods.apache;

import java.util.HashMap;
import java.util.Map;

import org.odftoolkit.odfdom.dom.style.props.OdfStyleProperty;
import org.odftoolkit.odfdom.dom.style.props.OdfTextProperties;

import com.github.jferard.spreadsheetwrapper.Util;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;

final class OdsOdfdomStyleFontHelper {
	private OdsOdfdomStyleFontHelper() {}

	public final static Map<OdfStyleProperty, String> getFontProperties(
			final WrapperFont wrapperFont) {
		final Map<OdfStyleProperty, String> fontProperties = new HashMap<OdfStyleProperty, String>(); // NOPMD
		final int bold = wrapperFont.getBold();
		final double size = wrapperFont.getSize();
		final int italic = wrapperFont.getItalic();
		final WrapperColor fontColor = wrapperFont.getColor();
		final String fontFamily = wrapperFont.getFamily();
		if (bold == WrapperCellStyle.YES) {
			OdsOdfdomStyleFontHelper.setProperties(fontProperties,
					OdsConstants.BOLD_ATTR_VALUE, OdfTextProperties.FontWeight,
					OdfTextProperties.FontWeightAsian,
					OdfTextProperties.FontWeightComplex);
		} else if (bold == WrapperCellStyle.NO) {
			OdsOdfdomStyleFontHelper.setProperties(fontProperties,
					OdsConstants.NORMAL_ATTR_VALUE,
					OdfTextProperties.FontWeight,
					OdfTextProperties.FontWeightAsian,
					OdfTextProperties.FontWeightComplex);
		}
	
		if (italic == WrapperCellStyle.YES) {
			OdsOdfdomStyleFontHelper.setProperties(fontProperties,
					OdsConstants.ITALIC_ATTR_VALUE,
					OdfTextProperties.FontStyle,
					OdfTextProperties.FontStyleAsian,
					OdfTextProperties.FontStyleComplex);
		} else if (italic == WrapperCellStyle.NO) {
			OdsOdfdomStyleFontHelper.setProperties(fontProperties,
					OdsConstants.NORMAL_ATTR_VALUE,
					OdfTextProperties.FontStyle,
					OdfTextProperties.FontStyleAsian,
					OdfTextProperties.FontStyleComplex);
		}
	
		if (!Util.almostEqual(size, WrapperCellStyle.DEFAULT)) {
			fontProperties.put(OdfTextProperties.FontSize,
					Double.toString(size) + "pt");
		}
	
		if (fontColor != null) {
			fontProperties.put(OdfTextProperties.Color, fontColor.toHex());
		}
	
		if (fontFamily != null) {
			OdsOdfdomStyleFontHelper.setProperties(fontProperties, fontFamily,
					OdfTextProperties.FontFamily, OdfTextProperties.FontName,
					OdfTextProperties.FontFamilyAsian,
					OdfTextProperties.FontNameAsian,
					OdfTextProperties.FontFamilyComplex,
					OdfTextProperties.FontNameComplex);
		}
		return fontProperties;
	}

	private static void setProperties(
			final Map<OdfStyleProperty, String> properties,
			final String attributeValue,
			final OdfStyleProperty... propertyArray) {
		for (final OdfStyleProperty property : propertyArray)
			properties.put(property, attributeValue);
	}

	public static WrapperFont toWrapperCellFont(
			final Map<OdfStyleProperty, String> propertyByAttrName) {
		final WrapperFont wrapperFont = new WrapperFont();
		final String fontWeight = propertyByAttrName
				.get(OdfTextProperties.FontWeight);
		final String fontStyle = propertyByAttrName
				.get(OdfTextProperties.FontStyle);
		final String fontSize = propertyByAttrName
				.get(OdfTextProperties.FontSize);
		final String fontColor = propertyByAttrName
				.get(OdfTextProperties.Color);
		final String fontFamily = propertyByAttrName
				.get(OdfTextProperties.FontFamily);
		if (fontWeight != null) {
			if (fontWeight.equals("bold"))
				wrapperFont.setBold();
			else if (fontWeight.equals(OdsConstants.NORMAL_ATTR_VALUE))
				wrapperFont.setBold(WrapperCellStyle.NO);
		}
		if (fontSize != null)
			wrapperFont.setSize(OdsConstants.sizeToPoints(fontSize));
		if (fontColor != null)
			wrapperFont.setColor(WrapperColor.stringToColor(fontColor));
		if (fontFamily != null)
			wrapperFont.setFamily(fontFamily);
		if (fontStyle != null) {
			if (fontStyle.equals(OdsConstants.ITALIC_ATTR_VALUE))
				wrapperFont.setItalic();
			else if (fontStyle.equals(OdsConstants.NORMAL_ATTR_VALUE))
				wrapperFont.setItalic(WrapperCellStyle.NO);
		}
		return wrapperFont;
	}

}
