package com.openxcell.util;

import java.util.Objects;

/**
 * @author vicky.thakor
 * @since 2018-05-15
 */
public class StringUtils {
	
	/**
	 * @param value
	 * @return
	 */
	public static boolean nonNullNotEmpty(String value) {
		return Objects.nonNull(value) && !value.trim().isEmpty();
	}
	
	/**
	 * @param value
	 * @param message
	 * @return
	 * @throws RuntimeException
	 */
	public static boolean requireNonNullNotEmpty(String value, String message) {
		if(!nonNullNotEmpty(value)) {
			throw new RuntimeException(message);
		}
		return true;
	}
}
