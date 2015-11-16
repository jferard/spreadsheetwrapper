package com.github.jferard.spreadsheetwrapper;

/*>>> import org.checkerframework.checker.nullness.qual.Nullable;*/

public class Util {
	public static boolean equal(final /*@Nullable*/ Object o, final /*@Nullable*/ Object p) {
		return o == null ? p == null : o.equals(p);
	}

	public static int hash(final /*@Nullable*/ Object o) {
		return o == null ? 0 : o.hashCode();
	}
}
