package com.openxcell.writer.spreadsheet;

import java.lang.reflect.InvocationTargetException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.json.JSONException;
import org.json.JSONObject;

import com.openxcell.util.BeanPropertyHolder;
import com.openxcell.util.BeanPropertyHolder.DataType;
import com.openxcell.writer.spreadsheet.SpreadSheetUtil.ExcelCellType;

/**
 * @author vicky.thakor
 * @since 2018-05-15
 */
public class SpreadSheetBeanUtil {
	/* Object caching */
	Map<Object, Map<String, BeanPropertyHolder>> mapObjectCache;

	String regExMatchPositionValue = "[a-zA-Z]*\\[([0-9]*?)\\]";

	String regExPositionExpression = "\\[([0-9]*?)\\]";

	/**
	 * Pattern used to get index from String. i.e: We want to access only 2nd index
	 * of Set<> or List<> then we pass String like follow. Pattern used to get value
	 * `2` from below String.
	 * 
	 * Item.itemPrice[2].priceListId
	 */
	Pattern pattern = Pattern.compile(regExPositionExpression);

	String regExMatchExpressionValue = "[a-zA-Z_]*\\(([a-zA-Z=0-9 \"\';!]*?)\\)";

	String regExStringExpresions = "\\(([a-zA-Z=0-9 \"\';!_]*?)\\)";

	/**
	 * Pattern : Pattern used to validate the expression. i.e: We want value only if
	 * it matches the given expression at last. Only return value when for that
	 * ItemPrice > priceListId=2
	 * 
	 * Item.itemPrice.value(priceListId=2)
	 */
	Pattern patternExpression = Pattern.compile(regExStringExpresions);

	private SpreadSheetTemplate excelCriteria;

	private SpreadSheetManager excelManager;

	/* Preserve original column property value (i.e Item.itemStore.store.name) */
	private String originalColumnProperty;

	public SpreadSheetBeanUtil() {
		mapObjectCache = new HashMap<>();
	}

	public void setOriginalColumnProperty(String originalColumnProperty) {
		this.originalColumnProperty = originalColumnProperty;
	}

	public void usingExcelCriteria(SpreadSheetTemplate excelCriteria) {
		this.excelCriteria = excelCriteria;
	}

	public void usingExcelManager(SpreadSheetManager excelManager) {
		this.excelManager = excelManager;
	}

	/**
	 * Evaluate column value recursively. Get nth level child value by performing
	 * substring upto nth period(`.`).<br/>
	 * <br/>
	 * Example:
	 * <ul>
	 * <li>Item.itemWarehouse.warehouseName</li>
	 * <li>itemWarehouse.warehouseName</li>
	 * <li>warehouseName</li>
	 * </ul>
	 * 
	 * @author vicky.thakor
	 * @date 15th April, 2015
	 * @param columnProperty
	 * @param object
	 * @return
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 * @throws NoSuchMethodException
	 * @throws SecurityException
	 * @throws JSONException
	 */
	public Object evaluateColumnValueRecursive(String columnProperty, Object object)
			throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException,
			SecurityException, JSONException {
		/* Where dealing with collection of data */
		String columnName;
		/* Default value of column when no value found */
		Object columnValue = "";

		/* Get Map<"column_name", ConfiguredColumn> */
		Map<String, BeanPropertyHolder> mapConfiguredColumn = new HashMap<>();

		if (columnProperty != null && !columnProperty.isEmpty() && object != null) {
			/* Check if period (`.`) exists in String or not */
			int indexOfExistPeriod = columnProperty.indexOf(".");

			if (indexOfExistPeriod > 0) {
				/**
				 * Split Object and Column name in Array i.e: Item.itemWarehouse.warehouseName
				 * -> [Item, itemWarehouse, warehouseName]
				 */
				String[] arrayProperties = columnProperty.split("\\.");
				/**
				 * SubString `strProperty` i.e: Item.itemWarehouse.warehouseName
				 * itemWarehouse.warehouseName warehouseName
				 */
				columnProperty = columnProperty.substring(indexOfExistPeriod + 1, columnProperty.length());
				/**
				 * Caching mechanism Reflection is bit costly in java so we created Map<Integer,
				 * Map<String, ConfiguredColumn>> to hold the objects previously loaded by
				 * reflection.
				 * 
				 * Example: "Item.sku","Item.name" Here we are accessing two properties of same
				 * object so when `Item` object loaded by reflection for `Item.sku` we store it
				 * in Map. Now when we want to access `Item.name` first check that object is
				 * available in Map or not if its available we don't need to perform reflection.
				 */
				try {
					/* @since 2016-11-18 mapObjectCache => mapKey changed from hashKey to object */
					if (object != null && mapObjectCache.containsKey(object)) {
						mapConfiguredColumn = mapObjectCache.get(object);
					} else {
						mapConfiguredColumn = new BeanPropertyHolder().makeReflactionCall(object);
						// if(object.hashCode() != 0){
						mapObjectCache.put(object, mapConfiguredColumn);
						// }
					}
				} catch (Exception e) {
					e.printStackTrace();
					columnValue = "#Error";
				}
				/**
				 * We are attaching expressions in String to get the specific value from SET<>
				 * or LIST<> so we are replacing expression with blank to get original column
				 * name.
				 * 
				 * .replaceAll("\\[[0-9]*\\]", "") for `Pattern 1` .replaceAll("\\([a-zA-Z=0-9
				 * \"\';]*\\)", "") for `Pattern 2`
				 * ========================================================== Pattern 1: Pattern
				 * used to get index from String. i.e: We want to access only 2nd index of Set<>
				 * or List<> then we pass String like follow. Pattern used to get value `2` from
				 * below String.
				 * 
				 * Item.itemPrice[2].priceListId
				 * ========================================================== Pattern 2: Pattern
				 * used to validate the expression. i.e: We want value only if it matches the
				 * given expression at last. Only return value when for that ItemPrice =>
				 * priceListId=2
				 * 
				 * Item.itemPrice.value(priceListId=2)
				 */
				String configuredColumnName = arrayProperties[1].replaceAll(regExPositionExpression, "")
						.replaceAll(regExStringExpresions, "");

				/* Get object of ConfiguredColumn from mapConfiguredColumn */
				BeanPropertyHolder objConfiguredColumn = mapConfiguredColumn.get(configuredColumnName);
				if (objConfiguredColumn != null) {
					/**
					 * Based on type of column perform operations. i.e: - For Set and List, loop the
					 * Objects. - For Integer, String, Date, etc... return the value.
					 */
					if (DataType.SET == objConfiguredColumn.getDataType()) {
						@SuppressWarnings("unchecked")
						Set<Object> setObject = (Set<Object>) objConfiguredColumn.getValue();
						if (setObject != null && !setObject.isEmpty()) {
							/**
							 * Case: When Set<> contains one Object return the value
							 * 
							 * Case: When Set<> contains more than one Object return value in following
							 * format: {value1; value2; ...}
							 */
							if (setObject.size() == 1) {
								for (Object getObject : setObject) {
									columnName = collectionDataHeader(columnProperty, getObject);
									columnValue = evaluateColumnValueRecursive(columnProperty, getObject);

									/* Multi-Warehouse changes 28/03/2016 */
									columnValue = collectionColumnNameColumnValue(columnName, columnValue);
								}
							} else if (arrayProperties.length > 2
									&& arrayProperties[1].matches(regExMatchPositionValue)) {
								int indexPosition = 0;
								String expressionValue = getExpressionBody(pattern, arrayProperties[1]);
								if (expressionValue != null && !expressionValue.isEmpty()) {
									indexPosition = Integer.valueOf(expressionValue);
								}
								indexPosition = indexPosition != 0 ? indexPosition - 1 : indexPosition;
								int setCount = 0;
								for (Object getObject : setObject) {
									if (setCount == indexPosition) {
										columnName = collectionDataHeader(columnProperty, getObject);
										columnValue = evaluateColumnValueRecursive(columnProperty, getObject);

										/* Multi-Warehouse changes 28/03/2016 */
										columnValue = collectionColumnNameColumnValue(columnName, columnValue);
									}
									setCount++;
								}
							} else {
								Object returnValue = null;
								for (Object getObject : setObject) {
									columnName = collectionDataHeader(columnProperty, getObject);
									returnValue = evaluateColumnValueRecursive(columnProperty, getObject);

									/* Multi-Warehouse changes 28/03/2016 */
									returnValue = collectionColumnNameColumnValue(columnName, returnValue);

									if (returnValue != null && !String.valueOf(returnValue).isEmpty()) {
										if (!"#ColumnValue".equalsIgnoreCase(String.valueOf(returnValue))) {
											columnValue += String.valueOf(returnValue) + ";";
										} else {
											columnValue = returnValue;
										}
									}
								}
								/**
								 * When we get multiple values from Set<?> then add `{....}`
								 */
								int semicolonCount = countCharacters(String.valueOf(columnValue), ';');
								if (columnValue != null && !String.valueOf(columnValue).isEmpty()
										&& semicolonCount > 1) {
//									columnValue = "{" + columnValue + "}";
									columnValue = String.valueOf(columnValue).split(";");
								} else if (semicolonCount == 1) {
									columnValue = String.valueOf(columnValue).replace("; ", "");
								}
							}
						}
					} else if (DataType.LIST == objConfiguredColumn.getDataType()) {
						@SuppressWarnings("unchecked")
						List<Object> listObject = (List<Object>) objConfiguredColumn.getValue();
						if (listObject != null && !listObject.isEmpty()) {
							/**
							 * Case: When List<> contains one Object return the value
							 * 
							 * Case: When List<> contains more than one Object return value in following
							 * format: {value1; value2; ...}
							 */
							if (listObject.size() == 1) {
								columnName = collectionDataHeader(columnProperty, listObject.get(0));
								columnValue = evaluateColumnValueRecursive(columnProperty, listObject.get(0));

								/* Multi-Warehouse changes 28/03/2016 */
								columnValue = collectionColumnNameColumnValue(columnName, columnValue);
							} else if (arrayProperties.length > 1
									&& arrayProperties[1].matches(regExMatchPositionValue)) {
								int indexPosition = 0;
								String expressionValue = getExpressionBody(pattern, arrayProperties[1]);
								if (expressionValue != null && !expressionValue.isEmpty()) {
									indexPosition = Integer.valueOf(expressionValue);
								}
								indexPosition = indexPosition != 0 ? indexPosition - 1 : indexPosition;

								columnName = collectionDataHeader(columnProperty, listObject.get(indexPosition));
								columnValue = evaluateColumnValueRecursive(columnProperty,
										listObject.get(indexPosition));

								/* Multi-Warehouse changes 28/03/2016 */
								columnValue = collectionColumnNameColumnValue(columnName, columnValue);
							} else {
								Object returnValue = null;
								for (Object getObject : listObject) {
									columnName = collectionDataHeader(columnProperty, getObject);
									returnValue = evaluateColumnValueRecursive(columnProperty, getObject);

									/* Multi-Warehouse changes 28/03/2016 */
									returnValue = collectionColumnNameColumnValue(columnName, returnValue);

									if (returnValue != null && !String.valueOf(returnValue).isEmpty()) {
										if (!"#ColumnValue".equalsIgnoreCase(String.valueOf(returnValue))) {
											columnValue += String.valueOf(returnValue) + ";";
										} else {
											columnValue = returnValue;
										}
									}
								}
								/**
								 * When we get multiple values from List<?> then add `{....}`
								 */
								int semicolonCount = countCharacters(String.valueOf(columnValue), ';');
								if (columnValue != null && !String.valueOf(columnValue).isEmpty()
										&& semicolonCount > 1) {
//									columnValue = "{" + columnValue + "}";
									columnValue = String.valueOf(columnValue).split(";");
								} else if (semicolonCount == 1) {
									columnValue = String.valueOf(columnValue).replace("; ", "");
								}
							}
						}
					} else if (DataType.OBJECT == objConfiguredColumn.getDataType()) {
						columnValue = evaluateColumnValueRecursive(columnProperty,
								objConfiguredColumn.getValue());
					} else {
						if (columnProperty.matches(regExMatchExpressionValue)) {
							boolean expressionCriteriaMatched = true;
							String expressionValue = getExpressionBody(patternExpression, columnProperty);
							/**
							 * batchExpression: priceListId=2;status!="Active" Reference:
							 * com.org.openxcell.report > InventoryStatus.json
							 */
							String[] batchExpression = expressionValue.split(";");
							for (String expression : batchExpression) {
								String[] validationExpression = null;
								BeanPropertyHolder configuredColumnExpression;
								String operationType = "";

								if (expression.contains("!=")) {
									validationExpression = expression.split("!=");
									operationType = "NOT_EQUAL";
								} else if (expression.contains("=")) {
									validationExpression = expression.split("=");
									operationType = "EQUAL";
								}
								if (validationExpression != null && validationExpression.length == 2) {
									configuredColumnExpression = mapConfiguredColumn.get(validationExpression[0]);
									if (configuredColumnExpression != null) {
										if (DataType.INTEGER == configuredColumnExpression.getDataType()) {
											if ("EQUAL".equals(operationType) && !(Integer
													.valueOf(
															String.valueOf(configuredColumnExpression.getValue()))
													.equals(Integer.valueOf(validationExpression[1])))) {
												expressionCriteriaMatched = false;
											} else if ("NOT_EQUAL".equals(operationType) && !(Integer.valueOf(String
													.valueOf(configuredColumnExpression.getValue())) != Integer
															.valueOf(validationExpression[1]))) {
												expressionCriteriaMatched = false;
											}
										} else if (DataType.TEXT == configuredColumnExpression.getDataType()) {
											if ("EQUAL".equals(operationType)
													&& !(String.valueOf(configuredColumnExpression.getValue())
															.equals(validationExpression[1]))) {
												expressionCriteriaMatched = false;
											} else if ("EQUAL".equals(operationType)
													&& (String.valueOf(configuredColumnExpression.getValue())
															.equals(validationExpression[1]))) {
												expressionCriteriaMatched = false;
											}
										}
									}
								}
							}
							if (expressionCriteriaMatched) {
								columnValue = objConfiguredColumn.getValue();
							}
						} else {
							columnValue = objConfiguredColumn.getValue();
						}
					}

					columnValue = doReplace(originalColumnProperty, columnValue);
					columnValue = doFormat(originalColumnProperty, columnValue);
				}
			}
		}
		return columnValue;
	}

	/**
	 * Get body part of given expression.<br/>
	 * Example:<br/>
	 * <b>1</b>. itemPrice[2] -> {Match `[Number]` Expression} -> return `2`<br/>
	 * <b>2</b>. value(priceListId=2) -> {Match `(property=value)` expression} ->
	 * return `priceListId=2` <br/>
	 * <br/>
	 * 
	 * @author vicky.thakor
	 * @param matchPattern
	 * @param str
	 * @return
	 */
	private String getExpressionBody(Pattern matchPattern, String str) {
		String strBody = "";
		Matcher matcher = matchPattern.matcher(str);
		if (matcher.find()) {
			strBody = matcher.group(1);
		}
		return strBody;
	}

	/**
	 * Count the occurrences of character in String.<br/>
	 * <br/>
	 * .
	 * 
	 * @author vicky.thakor
	 * @param str
	 * @param character
	 * @return
	 */
	private int countCharacters(String str, char character) {
		int counter = 0;
		if (str != null && !str.isEmpty()) {
			for (int i = 0; i < str.length(); i++) {
				if (str.charAt(i) == character) {
					counter++;
				}
			}
		}
		return counter;
	}

	/**
	 * When we've request to replace final value with predefined value. Example:
	 * When value of `orderStatus` is 5 then return `Shipped` as value for excel.
	 * "Shipment.salesOrder.orderStatus" : { "1":"Not Confirmed" "2":"Confirmed"
	 * "3":"Packed" "4":"Partial Shipped" "5":"Shipped" "6":"Delivered"
	 * "7":"Canceled" "8":"Return" }
	 *
	 * Reference: com.org.openxcell.report > ShipmentSummary.json
	 * 
	 * @param columnProperty
	 * @param columnValue
	 * @return
	 */
	private Object doReplace(String columnProperty, Object columnValue) {
		if (Objects.nonNull(excelCriteria.getReplace()) && excelCriteria.getReplace().has(columnProperty)) {
			JSONObject replaceValues = excelCriteria.getReplace().optJSONObject(columnProperty);

			if (replaceValues == null) {
				/* When calculating value of columnName in collectionDataFormat. */
				excelCriteria.getReplace().optJSONObject(columnProperty);
			}

			if (Objects.nonNull(replaceValues) && replaceValues.has(String.valueOf(columnValue))) {
				columnValue = replaceValues.opt(String.valueOf(columnValue));
			}
		}
		return columnValue;
	}

	/**
	 * Change format of value. "formatValue": { "Item.id": "TEXT" }
	 * 
	 * @param columnProperty
	 * @param columnValue
	 * @return
	 */
	private Object doFormat(String columnProperty, Object columnValue) {
		if (!String.valueOf(columnValue).trim().isEmpty()) {
			if (Objects.nonNull(excelCriteria.getFormatCellValue()) && excelCriteria.getFormatCellValue().has(columnProperty)) {
				String formatType = excelCriteria.getFormatCellValue().optString(columnProperty);
				if (ExcelCellType.INTEGER.toString().equalsIgnoreCase(formatType)) {
					columnValue = Integer.valueOf(String.valueOf(columnValue));
				} else if (ExcelCellType.FLOAT.toString().equalsIgnoreCase(formatType)) {
					columnValue = Double.valueOf(String.valueOf(columnValue));
				} else if (ExcelCellType.TEXT.toString().equalsIgnoreCase(formatType)) {
					columnValue = String.valueOf(columnValue);
				}
			}
		}
		return columnValue;
	}

	/**
	 * Get column name from data in case of Collection of data.
	 * TODO Will release in next version
	 * 
	 * @param object
	 * @return
	 */
	private String collectionDataHeader(String columnProperty, Object object) {
		String columnName = "#ColumnName";
//		if (Objects.nonNull(excelCriteria.getCollectionDataFormat())) {
//			JSONObject collectionDataFormat = excelCriteria.getCollectionDataFormat().optJSONObject(columnProperty);
//			if (Objects.nonNull(collectionDataFormat)) {
//				columnName = collectionDataFormat.optString("columnName");
//				String prefix = collectionDataFormat.optString("prefix");
//				String suffix = collectionDataFormat.optString("suffix");
//				if (columnName.contains(".")) {
//					try {
//						columnName = String.valueOf(evaluateColumnValueRecursive(columnName, object));
//						columnName = String
//								.valueOf(doReplace(collectionDataFormat.optString("columnName"), columnName));
//						if (!prefix.isEmpty()) {
//							columnName = prefix + columnName;
//						}
//						if (!suffix.isEmpty()) {
//							columnName += suffix;
//						}
//
//					} catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException
//							| NoSuchMethodException | SecurityException | JSONException e) {
//						e.printStackTrace();
//					}
//				}
//			}
//		}
		return columnName;
	}

	/**
	 * Create header in excel for collection of data.
	 * 
	 * @param columnName
	 * @param columnValue
	 * @return
	 */
	private Object collectionColumnNameColumnValue(String columnName, Object columnValue) {
		/* Multi-Warehouse changes 28/03/2016 */
		if (!"#ColumnName".equalsIgnoreCase(columnName)) {
			excelManager.addHeader(columnName);
			excelManager.addValueCell(columnName, columnValue);
			columnValue = "#ColumnValue";
		}
		return columnValue;
	}
}
