package com.github.jferard.spreadsheetwrapper.ods.apache;

import java.util.HashMap;
import java.util.Map;

import org.odftoolkit.odfdom.dom.style.props.OdfStyleProperty;
import org.odftoolkit.odfdom.dom.style.props.OdfTableCellProperties;

import com.github.jferard.spreadsheetwrapper.ods.OdsConstants;
import com.github.jferard.spreadsheetwrapper.style.Borders;
import com.github.jferard.spreadsheetwrapper.style.WrapperCellStyle;
import com.github.jferard.spreadsheetwrapper.style.WrapperColor;

final class OdsApacheStyleBorderHelper {
	private OdsApacheStyleBorderHelper() {}

	public static Borders toCellBorders(
			final Map<OdfStyleProperty, String> propertyByAttrName) {
		final Borders borders = new Borders();
		final String borderAttrValue = propertyByAttrName
				.get(OdfTableCellProperties.Border);
		if (borderAttrValue != null) {
			final String[] split = borderAttrValue.split("\\s+");
			borders.setLineWidth(OdsConstants.sizeToPoints(split[0]));
			// borders.setLineType(split[1]);
			if (!WrapperColor.BLACK.toHex().equals(split[2]))
				borders.setLineColor(WrapperColor.stringToColor(split[2]));
		}
		
		// Here : specify each border
		return borders;
	}

	public static Map<OdfStyleProperty, String> getBordersProperties(
			final Borders borders) {
		final Map<OdfStyleProperty, String> bordersProperties = new HashMap<OdfStyleProperty, String>(); // NOPMD
		StringBuilder builder = new StringBuilder();
		final double borderLineWidth = borders.getLineWidth();
		if (borderLineWidth > 0.0) {
			builder.append(borderLineWidth).append("pt");
			builder.append(" solid ");
			final WrapperColor borderLineColor = borders.getLineColor();
			if (borderLineColor == null)
				builder.append("#000000");
			else
				builder.append(borderLineColor.toHex());
			bordersProperties.put(OdfTableCellProperties.Border,
					builder.toString());
		}
		// Here : specify each border
		return bordersProperties;
	}

}
