package com.openxcell.util;

import java.util.Objects;

/**
 * @author vicky.thakor
 * @since 2018-05-17
 */
public class RegExUtils {
	/**
	 * Verify the given string is number only (i.e: 1, 22, 45, -1, -22, -45 etc...)
	 *
	 * @param value
	 * @return
	 */
	public static boolean isNumber(String value) {
		return value != null && value.matches("^-?[0-9]\\d*(\\.\\d+)?$");
	}

	/**
	 * Verify the given string is boolean
	 *
	 * @param value
	 * @return
	 */
	public static boolean isBoolean(String value) {
		return value != null && value.matches("^(?i)(true|false)$");
	}

	/**
	 * Verify the given string is alpha-numeric if value is empty or null then false
	 * will be return - "a" : pass - "a123" : pass - "452": pass - "a#23": fail - "a
	 * 45": fail
	 *
	 * @param value
	 * @return
	 */
	public static boolean isAlphaNumeric(String value) {
		return (value != null && !value.isEmpty()) && value.matches("^[a-zA-Z0-9]*$");
	}

	/**
	 * Verify the given string is number / floating point value only (i.e: 1, 22.32,
	 * 45, etc...)
	 *
	 * @param value
	 * @return
	 */
	public static boolean isFloatingNumber(String value) {
		return value != null && value.matches("[0-9]*\\.?[0-9]+");
	}
	
	/**
	 * Verify the given string is number (1, -2, 3, etc...)
	 * @param value
	 * @return
	 */
	public static boolean isInteger(String value) {
		return Objects.nonNull(value) && value.matches("^-?[0-9]+");
	}
}
