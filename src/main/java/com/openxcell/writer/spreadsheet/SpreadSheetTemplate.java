package com.openxcell.writer.spreadsheet;

import org.json.JSONArray;
import org.json.JSONObject;

import com.openxcell.util.StringUtils;

/**
 * @author vicky.thakor
 * @since 2018-05-15
 */
public class SpreadSheetTemplate {
	private JSONArray header;
	private JSONObject headerTextColor;
	private JSONObject headerBackgroundColor;
	private JSONArray properties;
	private JSONArray summationHeaders;

	private JSONObject replace;
	private JSONObject extendedReplace;
	
	private JSONObject formatCellValue;

	public SpreadSheetTemplate(String jsonTemplate) {
		StringUtils.requireNonNullNotEmpty(jsonTemplate, "template can not be null or empty");
		try {
			JSONObject jsonObject = new JSONObject(jsonTemplate);
			header = jsonObject.getJSONArray("header");
			headerTextColor = jsonObject.optJSONObject("header_text_color");
			headerBackgroundColor = jsonObject.optJSONObject("header_background_color");
			properties = jsonObject.optJSONArray("properties");
			summationHeaders = jsonObject.optJSONArray("summation_headers");

			replace = jsonObject.optJSONObject("replace");
			extendedReplace = jsonObject.optJSONObject("extended_replace");
			
			formatCellValue = jsonObject.optJSONObject("format_cell_value");
		} catch (Exception e) {
			throw e;
		}
	}

	public JSONArray getHeader() {
		return header;
	}

	public JSONObject getHeaderTextColor() {
		return headerTextColor;
	}

	public JSONObject getHeaderBackgroundColor() {
		return headerBackgroundColor;
	}

	public JSONArray getProperties() {
		return properties;
	}

	public JSONArray getSummationHeaders() {
		return summationHeaders;
	}
	
	public JSONObject getReplace() {
		return replace;
	}
	
	public JSONObject getExtendedReplace() {
		return extendedReplace;
	}
	
	public JSONObject getFormatCellValue() {
		return formatCellValue;
	}
}