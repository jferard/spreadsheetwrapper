package com.github.jferard.spreadsheetwrapper;

public class Util {
	public static boolean equal(final Object o, final Object p) {
		return o == null ? p == null : o.equals(p);
	}

	public static int hash(final Object o) {
		return o == null ? 0 : o.hashCode();
	}
}
