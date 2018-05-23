package com.openxcell.writer.spreadsheet;

import java.lang.reflect.InvocationTargetException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import com.openxcell.util.CollectionUtils;
import com.openxcell.util.JSONUtils;
import com.openxcell.util.StringUtils;

/**
 * @author vicky.thakor
 * @since 2018-05-15
 */
public class SpreadSheetBeanManager extends SpreadSheetManager {
	private static Logger logger = Logger.getLogger(SpreadSheetBeanManager.class.getName());
	private SpreadSheetTemplate sheetTemplate;
	private SpreadSheetBeanUtil beanUtil;

	private boolean isExtendedReplace = false;
	private Map<String, Object> mapExtendedReplaceProperties = new HashMap<String, Object>(0);

	public SpreadSheetBeanManager(SpreadSheetTemplate sheetTemplate) {
		this.sheetTemplate = sheetTemplate;
		buildWorkbook();
		freezeHeader();
		boldHeader();
		init();
	}

	private void init() {
		beanUtil = new SpreadSheetBeanUtil();
		beanUtil.usingExcelCriteria(sheetTemplate);
		beanUtil.usingExcelManager(this);
	}

	public boolean process(List<?> listData) {
		summationProperties();
		headerTextColor();
		headerBackgroundColor();

		if (Objects.nonNull(sheetTemplate.getHeader())) {
			for (int i = 0; i < sheetTemplate.getHeader().length(); i++) {
				addHeader(sheetTemplate.getHeader().optString(i).trim());
			}
		}

		if (Objects.nonNull(sheetTemplate.getProperties()) && Objects.nonNull(listData)) {
			int columnCount = 0;
			for (Object object : listData) {
				isExtendedReplace = false;
				columnCount++;
				newRow();
				Map<Integer, String[]> collectionDataCache = new HashMap<>();
				
				for (int i = 0; i < sheetTemplate.getProperties().length(); i++) {
					String columnProperty = sheetTemplate.getProperties().optString(i).trim();

					if (Objects.nonNull(columnProperty) && !columnProperty.isEmpty()) {
						if ("count".equalsIgnoreCase(columnProperty)) {
							addValueCell(columnCount);
						} else {
							try {
								Object columnValue;
								if (isExtendedReplace && mapExtendedReplaceProperties.containsKey(columnProperty)) {
									columnValue = mapExtendedReplaceProperties.get(columnProperty);
								} else {
									beanUtil.setOriginalColumnProperty(columnProperty);
									columnValue = beanUtil.evaluateColumnValueRecursive(columnProperty, object);
									columnValue = Objects.nonNull(columnValue) ? columnValue : "";
									processExtendedReplace(columnProperty, columnValue);
								}

								if (!"#ColumnValue".equalsIgnoreCase(String.valueOf(columnValue))) {
									if(columnValue instanceof String[]) {
										String[] values = (String[])columnValue;
										addValueCell(values[0]);
										collectionDataCache.put(i, values);
									}else {
										addValueCell(columnValue);
									}
								}
							} catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException
									| NoSuchMethodException | SecurityException | JSONException e) {
								logger.log(Level.SEVERE, e.getMessage(), e);
							}
						}
					}
				}
				
				if(CollectionUtils.nonNullNonEmptyMap(collectionDataCache)) {
					AtomicInteger newRowCount = new AtomicInteger(0);
					for(String[] values : collectionDataCache.values()) {
						/* length - 1 => 1st element from array is added in first iteration */
						int arraylength = values.length - 1;
						if(newRowCount.get() < arraylength) {
							newRowCount.set(arraylength);
						}
					}
					
					for (int i = 0; i < newRowCount.get(); i++) {
						newRow();
						for(Map.Entry<Integer, String[]> entrySet : collectionDataCache.entrySet()) {
							String[] values = entrySet.getValue();
							if(values.length > (i+1) && StringUtils.nonNullNotEmpty(values[i+1])) {
								addValueCell(entrySet.getKey(), values[i+1]);
							}
						}
					}
				}
				
			}
		}

		return true;
	}

	/**
	 * Attach summation column details to {@link ExcelManager}
	 */
	private void summationProperties() {
		if (Objects.nonNull(sheetTemplate.getSummationHeaders())) {
			summationBeforeNewSheet(JSONUtils.JSONArrayToList(String.class, sheetTemplate.getSummationHeaders()));
		}
	}

	/**
	 * Set header background color
	 */
	private void headerTextColor() {
		if (Objects.nonNull(sheetTemplate.getHeaderTextColor())) {
			Map<String, Short> headerTextColor = new HashMap<>(0);
			Iterator<String> iterator = sheetTemplate.getHeaderTextColor().keys();
			while (iterator.hasNext()) {
				String key = iterator.next();
				Short color = Short.valueOf(sheetTemplate.getHeaderTextColor().optString(key));
				headerTextColor.put(key, color);
			}
			setHeaderTextColor(headerTextColor);
		}
	}

	/**
	 * Set header background color
	 */
	private void headerBackgroundColor() {
		if (Objects.nonNull(sheetTemplate.getHeaderBackgroundColor())) {
			Map<String, Short> headerBackgroundColor = new HashMap<>(0);
			Iterator<String> iterator = sheetTemplate.getHeaderBackgroundColor().keys();
			while (iterator.hasNext()) {
				String key = iterator.next();
				Short color = Short.valueOf(sheetTemplate.getHeaderBackgroundColor().optString(key));
				headerBackgroundColor.put(key, color);
			}
			setHeaderBackgroundColor(headerBackgroundColor);
		}
	}

	/**
	 * "ExtendedReplace": { "Item.type": { "value": "Configurable", "replace": [
	 * {"Item.quantity": ""}, {"Item.threshold": ""} ] } }
	 *
	 * @param columnProperty
	 * @param columnValue
	 */
	private void processExtendedReplace(String columnProperty, Object columnValue) {
		if (Objects.nonNull(sheetTemplate.getExtendedReplace())
				&& sheetTemplate.getExtendedReplace().has(columnProperty)) {
			JSONObject extendedReplaceProperties = sheetTemplate.getExtendedReplace().optJSONObject(columnProperty);
			String extendedReplaceValue = extendedReplaceProperties.optString("value");
			if (extendedReplaceValue.equals(String.valueOf(columnValue))) {
				isExtendedReplace = true;
				JSONArray replaceProperties = extendedReplaceProperties.optJSONArray("replace");
				if (Objects.nonNull(replaceProperties)) {
					for (int j = 0; j < replaceProperties.length(); j++) {
						JSONObject replaceProperty = replaceProperties.optJSONObject(j);
						if (Objects.nonNull(replaceProperty)) {
							Iterator<String> keys = replaceProperty.keys();
							while (keys.hasNext()) {
								String key = keys.next();
								String value = replaceProperty.optString(key);
								mapExtendedReplaceProperties.put(key, value);
							}
						}
					}
				}
			}
		}
	}
}
