package com.openxcell.util;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.logging.Logger;

import com.openxcell.annotation.IgnoreField;
import com.openxcell.util.BeanPropertyHolder.DataType;
import com.openxcell.writer.iface.ReaderWriterBean;

/**
 * @author vicky.thakor
 * @since 2018-05-16
 */
public class BeanPropertyValueLoader {

	private static Logger logger = Logger.getLogger(BeanPropertyValueLoader.class.getName());
	private Object localObject;

	public BeanPropertyValueLoader(Object obj) {
		this.localObject = obj;
	}

	/**
	 * Get object values in form of Map<{@link String},
	 * {@link BeanPropertyHolder}>.<br/>
	 * <br/>
	 * 
	 * @author vicky.thakor
	 * @date 26th April, 2015
	 */
	public Map<String, BeanPropertyHolder> getBeanProperties() {
		Map<String, BeanPropertyHolder> mapConfiguredColumns = new HashMap<String, BeanPropertyHolder>();
		Field[] arrayFields = localObject.getClass().getDeclaredFields();
		Method[] methods = localObject.getClass().getDeclaredMethods();

		/* access super class properties(variables) */
		if (localObject.getClass().getSuperclass() != null
				&& localObject.getClass().getSuperclass() != ReaderWriterBean.class
				&& ReaderWriterBean.class.isAssignableFrom(localObject.getClass().getSuperclass())) {
			Field[] temporaryFields = arrayFields;
			Field[] superFields = localObject.getClass().getSuperclass().getDeclaredFields();
			int totalFields = temporaryFields.length + superFields.length;
			arrayFields = new Field[totalFields];

			System.arraycopy(temporaryFields, 0, arrayFields, 0, temporaryFields.length);
			System.arraycopy(superFields, 0, arrayFields, temporaryFields.length, superFields.length);

			Method[] temporary = methods;
			Method[] superMethods = localObject.getClass().getSuperclass().getDeclaredMethods();
			int totalMethods = temporary.length + superMethods.length;
			methods = new Method[totalMethods];

			System.arraycopy(temporary, 0, methods, 0, temporary.length);
			System.arraycopy(superMethods, 0, methods, temporary.length, superMethods.length);
		}

		for (int i = 0; i < arrayFields.length; i++) {
			if (!arrayFields[i].isAnnotationPresent(IgnoreField.class)) {
				String columnName = arrayFields[i].getName();
				Object columnValue = null;
				if (StringUtils.nonNullNotEmpty(columnName) && !columnName.startsWith("$")) {
					String getMethodName = columnName.replaceFirst(columnName.substring(0, 1),
							columnName.substring(0, 1).toUpperCase());

					try {
						for (Method method : methods) {
							if (("get" + getMethodName).equalsIgnoreCase(method.getName())
									|| ("is" + getMethodName).equalsIgnoreCase(method.getName())) {
								columnValue = method.invoke(localObject);
								break;
							}
						}
					} catch (Exception e) {
						logger.info(e.getMessage());
					}
					mapConfiguredColumns.put(columnName,
							new BeanPropertyHolder(columnName, columnValue, getVariableDataType(columnValue)));
				}
			}
		}
		return mapConfiguredColumns;
	}

	/**
	 * Identify column value type.<br/>
	 * <br/>
	 * 
	 * @author vicky.thakor
	 * @date 26th April, 2015
	 * @param value
	 * @return
	 */
	private DataType getVariableDataType(Object value) {
		if (value instanceof Float || value instanceof Double
				|| value instanceof BigDecimal) {
			return DataType.FLOAT;
		} else if (value instanceof Integer || value instanceof Long || value instanceof Short
				|| value instanceof BigInteger) {
			return DataType.INTEGER;
		} else if (value instanceof Date) {
			return DataType.DATE;
		} else if (value instanceof Boolean) {
			return DataType.BOOLEAN;
		} else if (value instanceof String) {
			return DataType.TEXT;
		} else if (value instanceof Set<?>) {
			return DataType.SET;
		} else if (value instanceof List<?>) {
			return DataType.LIST;
		} else if (value instanceof Map<?, ?>) {
			return DataType.MAP;
		} else {
			return DataType.OBJECT;
		}
	}
}
