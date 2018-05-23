package com.openxcell.util;

import java.util.Collection;
import java.util.Map;
import java.util.Objects;

/**
 * @author vicky.thakor
 * @since 2018-05-17
 */
public class CollectionUtils {
	/**
	 * Check given collection is not null and not empty.
	 * 
	 * @author vicky.thakor
	 * @since 2017-01-09
	 * @param collection
	 * @return
	 */
	public static boolean nonNullNonEmpty(@SuppressWarnings("rawtypes") Collection collection) {
		return Objects.nonNull(collection) && !collection.isEmpty();
	}
	
	/**
	 * Check given map is not null and not empty.
	 * 
	 * @author vicky.thakor
	 * @since 2017-03-21
	 * @param map
	 * @return
	 */
	public static boolean nonNullNonEmptyMap(@SuppressWarnings("rawtypes") Map map) {
		return Objects.nonNull(map) && !map.isEmpty();
	}
}
