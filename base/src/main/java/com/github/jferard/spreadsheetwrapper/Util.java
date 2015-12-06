package com.github.jferard.spreadsheetwrapper;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

/**
 * Tiny Util class.
 */
public final class Util {
	/**
	 * @param object
	 * @param otherObject
	 * @return true if object and otherObject are equal
	 */
	public static boolean equal(final/*@Nullable*/Object object,
			final/*@Nullable*/Object otherObject) {
		return object == null ? otherObject == null : object
				.equals(otherObject);
	}
	
	/**
	 * @param aDouble
	 * @param otherDouble
	 * @return true if aDouble and otherDouble are equal to 0.001
	 */
	public static boolean almostEqual(final double aDouble,
			final double otherDouble) {
		return Math.abs(aDouble - otherDouble) < 0.001;
	}
	

	/**
	 * @param object
	 *            the objet to hash
	 * @return 0 if object is null, else object.hashCode
	 */
	public static int hash(final/*@Nullable*/Object object) {
		return object == null ? 0 : object.hashCode();
	}

	private Util() {
	}
}
