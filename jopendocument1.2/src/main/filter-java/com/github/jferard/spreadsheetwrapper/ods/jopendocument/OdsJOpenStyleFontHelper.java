package com.github.jferard.spreadsheetwrapper.ods.${jopendocument.pkg};

import org.jdom.Element;

import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;
import com.github.jferard.spreadsheetwrapper.style.WrapperFont;

class OdsJOpenStyleFontHelper {
	private OdsJOpenStyleFontHelper() {}

	public static void setFont(final Element textProps, final WrapperFont cellFont) {
		OdsJOpenStyleFontHelper.setFontBold(textProps, cellFont.getBold());
		OdsJOpenStyleFontHelper.setFontItalic(textProps, cellFont.getItalic());
		OdsJOpenStyleFontHelper.setFontSize(textProps, cellFont.getSize());
		OdsJOpenStyleFontHelper.setFontColor(textProps, cellFont.getColor());
		OdsJOpenStyleFontHelper.setFontFamily(textProps, cellFont.getFamily());
	}

	private static void setFontBold(final Element textProps, final int key) {
		final String value;
		if (key == WrapperCellStyle.YES)
			value = OdsConstants.BOLD_ATTR_VALUE;
		else if (key == WrapperCellStyle.YES)
			value = OdsConstants.NORMAL_ATTR_VALUE;
		else
			value = null;
	
		if (value != null) {
			textProps.setAttribute(OdsConstants.FONT_WEIGHT_ATTR_NAME, value,
					OdsJOpenStyleHelper.FO_NS);
			textProps.setAttribute(OdsConstants.FONT_WEIGHT_ASIAN_ATTR_NAME, value,
					OdsJOpenStyleHelper.STYLE_NS);
			textProps.setAttribute(OdsConstants.FONT_WEIGHT_COMPLEX_ATTR_NAME, value,
					OdsJOpenStyleHelper.STYLE_NS);
		}
	}

	private static void setFontColor(final Element textProps, final WrapperColor color) {
		if (color != null) {
			textProps.setAttribute(OdsConstants.COLOR_ATTR_NAME, color.toHex(),
					OdsJOpenStyleHelper.FO_NS);
		}
	}

	private static void setFontFamily(final Element textProps, final String family) {
		if (family != null) {
			textProps.setAttribute(OdsConstants.FAMILY_ATTR_NAME, family,
					OdsJOpenStyleHelper.FO_NS);
			textProps.setAttribute(OdsConstants.FAMILY_ASIAN_ATTR_NAME, family,
					OdsJOpenStyleHelper.STYLE_NS);
			textProps.setAttribute(OdsConstants.FAMILY_COMPLEX_ATTR_NAME, family,
					OdsJOpenStyleHelper.STYLE_NS);
		}
	}

	private static void setFontItalic(final Element textProps, final int key) {
		final String value;
		if (key == WrapperCellStyle.YES)
			value = OdsConstants.ITALIC_ATTR_VALUE;
		else if (key == WrapperCellStyle.YES)
			value = OdsConstants.NORMAL_ATTR_VALUE;
		else
			value = null;
	
		if (value != null) {
			textProps.setAttribute(OdsConstants.FONT_STYLE_ATTR_NAME, value,
					OdsJOpenStyleHelper.FO_NS);
			textProps.setAttribute(OdsConstants.FONT_STYLE_ASIAN_ATTR_NAME, value,
					OdsJOpenStyleHelper.STYLE_NS);
			textProps.setAttribute(OdsConstants.FONT_STYLE_COMPLEX_ATTR_NAME, value,
					OdsJOpenStyleHelper.STYLE_NS);
		}
	}

	private static void setFontSize(final Element textProps, final double size) {
		if (size != WrapperCellStyle.DEFAULT) {
			textProps.setAttribute(OdsConstants.FONT_SIZE_ATTR_NAME, Double.toString(size)+"pt",
					OdsJOpenStyleHelper.FO_NS);
		}
	}

	public static WrapperFont toWrapperFont(final Element odfElement) {
		final WrapperFont wrapperFont = new WrapperFont();
		final String fontWeight = odfElement.getAttributeValue(
				OdsConstants.FONT_WEIGHT_ATTR_NAME, OdsJOpenStyleHelper.FO_NS);
		if (OdsConstants.BOLD_ATTR_VALUE.equals(fontWeight))
			wrapperFont.setBold();
		else if (OdsConstants.NORMAL_ATTR_VALUE.equals(fontWeight))
			wrapperFont.setBold(WrapperCellStyle.NO);
	
		final String fontStyle = odfElement.getAttributeValue(
				OdsConstants.FONT_STYLE_ATTR_NAME, OdsJOpenStyleHelper.FO_NS);
		if (OdsConstants.ITALIC_ATTR_VALUE.equals(fontStyle))
			wrapperFont.setItalic();
		else if (OdsConstants.NORMAL_ATTR_VALUE.equals(fontStyle))
			wrapperFont.setItalic(WrapperCellStyle.NO);
	
		final String fontSize = odfElement.getAttributeValue(
				OdsConstants.FONT_SIZE_ATTR_NAME, OdsJOpenStyleHelper.FO_NS);
		if (fontSize != null)
			wrapperFont.setSize(OdsConstants.sizeToPoints(fontSize));
	
		final String fColorAsHex = odfElement.getAttributeValue(
				OdsConstants.COLOR_ATTR_NAME, OdsJOpenStyleHelper.FO_NS);
		final WrapperColor fontColor = WrapperColor.stringToColor(fColorAsHex);
		if (fontColor != null)
			wrapperFont.setColor(fontColor);
	
		final String fontFamily =  odfElement.getAttributeValue(
				OdsConstants.FAMILY_ATTR_NAME, OdsJOpenStyleHelper.FO_NS);
		if (fontFamily != null)
			wrapperFont.setFamily(fontFamily);
		return wrapperFont;
	}

}
