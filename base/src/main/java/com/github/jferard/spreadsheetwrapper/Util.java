package com.github.jferard.spreadsheetwrapper;

public class Util {
	public static boolean equal(Object o, Object p) {
		return o == null ? p == null : o.equals(p);
	}
	
	public static int hash(Object o) {
		return o == null ? 0 : o.hashCode();
	}
}
