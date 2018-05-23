package com.openxcell.util;

import java.lang.reflect.InvocationTargetException;
import java.util.Map;

/**
 * Custom object to hold the Entity's ColumnName, ColumnValue and Its data
 * type.<br/>
 * <br/>
 * 
 * @author vicky.thakor
 * @date 9th April, 2015
 */
public class BeanPropertyHolder {

	public enum DataType {
		TEXT, INTEGER, FLOAT, DATE, BOOLEAN, OBJECT, SET, LIST, MAP
	}

	private String property;
	private Object value;
	private DataType dataType;

	public BeanPropertyHolder() {
	}

	public BeanPropertyHolder(String property, Object value, DataType columnType) {
		setProperty(property);
		setValue(value);
		setDataType(columnType);
	}

	/**
	 * Call getConfiguredColumns() method of bean using Reflection.
	 * 
	 * @author vicky.thakor
	 * @param obj
	 * @return
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 * @throws NoSuchMethodException
	 * @throws SecurityException
	 */
	public Map<String, BeanPropertyHolder> makeReflactionCall(Object obj) throws IllegalAccessException,
			IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException {
		BeanPropertyValueLoader objGenericClassValueLoader = new BeanPropertyValueLoader(obj);
		return objGenericClassValueLoader.getBeanProperties();
	}

	public String getProperty() {
		return property;
	}

	public void setProperty(String property) {
		this.property = property;
	}

	public Object getValue() {
		return value;
	}

	public void setValue(Object value) {
		this.value = value;
	}

	public DataType getDataType() {
		return dataType;
	}

	public void setDataType(DataType dataType) {
		this.dataType = dataType;
	}
}
