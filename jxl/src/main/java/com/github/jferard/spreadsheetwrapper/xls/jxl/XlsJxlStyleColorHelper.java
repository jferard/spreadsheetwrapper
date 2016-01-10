package com.github.jferard.spreadsheetwrapper.xls.jxl;

import java.util.HashMap;
import java.util.Map;

import jxl.format.Colour;
import jxl.format.RGB;

import com.github.jferard.spreadsheetwrapper.style.WrapperColor;

class XlsJxlStyleColorHelper {
	/** wrapper -> internal */
	private final Map<WrapperColor, Colour> jxlColorByWrapperColor;
	/** internal -> wrapper */
	private final Map<Colour, WrapperColor> wrapperColorByJxlColor;

	XlsJxlStyleColorHelper() {
		final WrapperColor[] colors = WrapperColor.values();
		this.jxlColorByWrapperColor = new HashMap<WrapperColor, Colour>(
				colors.length);
		this.wrapperColorByJxlColor = new HashMap<Colour, WrapperColor>(
				colors.length);
		for (final WrapperColor color : colors) {
			Colour jxlColor;
			try {
				final Class<?> jxlClazz = Class.forName("jxl.format.Colour");
				jxlColor = (Colour) jxlClazz.getDeclaredField(
						color.getSimpleName()).get(null);
				if (jxlColor != null) {
					this.jxlColorByWrapperColor.put(color, jxlColor);
					this.wrapperColorByJxlColor.put(jxlColor, color);
				}
			} catch (final ClassNotFoundException e) {
				jxlColor = null;
			} catch (final IllegalArgumentException e) {
				jxlColor = null;
			} catch (final IllegalAccessException e) {
				jxlColor = null;
			} catch (final NoSuchFieldException e) {
				jxlColor = null;
			} catch (final SecurityException e) {
				jxlColor = null;
			}
		}
	}

	/**
	 * @param wrapperColor
	 *            the wrapper color
	 * @return the internal color
	 */
	public /*@Nullable*/Colour toJxlColor(final WrapperColor wrapperColor) {
		return this.jxlColorByWrapperColor.get(wrapperColor);
	}

	static boolean colorEquals(final Colour color1, final Colour color2) {
		final RGB rgb1 = color1.getDefaultRGB();
		final RGB rgb2 = color2.getDefaultRGB();
		return rgb1.getRed() == rgb2.getRed()
				&& rgb1.getGreen() == rgb2.getGreen()
				&& rgb1.getBlue() == rgb2.getBlue();
	}

	/**
	 * @param jxlColor
	 *            internal color
	 * @return wrapper color
	 */
	public/*@Nullable*/WrapperColor toWrapperColor(final Colour jxlColor) {
		return this.wrapperColorByJxlColor.get(jxlColor);
	}

}
