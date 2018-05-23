package com.openxcell.util;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import org.json.JSONArray;

/**
 * @author vicky.thakor
 * @since 2018-05-15
 */
public class JSONUtils {

	/**
	 * Convert JSONArray to List.
	 *
	 * @author vicky.thakor
	 * @param type
	 *            Long.class, Integer.class, String.class
	 * @param jsonArray
	 * @return
	 */
	public static <T> List<T> JSONArrayToList(Class<T> type, JSONArray jsonArray) {
		List<T> list = new ArrayList<>(0);
		if (Objects.nonNull(type) && Objects.nonNull(jsonArray)) {
			if (type.equals(Long.class)) {
				for (int i = 0; i < jsonArray.length(); i++) {
					list.add(type.cast(jsonArray.optLong(i)));
				}
			} else if (type.equals(Integer.class)) {
				for (int i = 0; i < jsonArray.length(); i++) {
					list.add(type.cast(jsonArray.optInt(i)));
				}
			} else if (type.equals(String.class)) {
				for (int i = 0; i < jsonArray.length(); i++) {
					list.add(type.cast(jsonArray.optString(i)));
				}
			} else if (type.equals(Object.class)) {
				for (int i = 0; i < jsonArray.length(); i++) {
					list.add(type.cast(jsonArray.opt(i)));
				}
			} else {
				for (int i = 0; i < jsonArray.length(); i++) {
					list.add(type.cast(jsonArray.opt(i)));
				}
			}
		}
		return list;
	}
}
