package com.github.jferard.spreadsheetwrapper.ods.${jopendocument.pkg};

import org.jdom.Element;

import com.github.jferard.spreadsheetwrapper.Util;
import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;
import com.github.jferard.spreadsheetwrapper.style.Borders;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;

class OdsJOpenStyleBorderHelper {
	private OdsJOpenStyleBorderHelper() {}

	public static void setBorders(final Element tableCellProps, final Borders borders) {
		final double lineWidth = borders.getLineWidth();
		if (!Util.almostEqual(lineWidth, WrapperCellStyle.DEFAULT)) {
			WrapperColor color = borders.getLineColor();
			if (color == null)
				color = WrapperColor.BLACK;
			
			tableCellProps.setAttribute(OdsConstants.BORDER_ATTR_NAME,
					Double.toString(lineWidth) + "pt solid "+color.toHex(),
					OdsJOpenStyleHelper.FO_NS);
		}
	}

	public static Borders toBorders(final Element cellPropertiesElment) {
		final Borders borders = new Borders();
		final String borderAttrValue = cellPropertiesElment
				.getAttributeValue(OdsConstants.BORDER_ATTR_NAME,
						OdsJOpenStyleHelper.FO_NS);
		if (borderAttrValue == null) {
			// border top
			// border bottom
			// ...
			return null;
		} else {
			final String[] split = borderAttrValue.split("\\s+");
			borders.setLineWidth(OdsConstants.sizeToPoints(split[0]));
			// borders.setLineType(OdsConstants.sizeToPoints(split[1]));
			if (!WrapperColor.BLACK.toHex().equals(split[2]))
				borders.setLineColor(WrapperColor.stringToColor(split[2]));
		}
		return borders;
	}

}
