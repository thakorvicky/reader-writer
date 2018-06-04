package com.openxcell.writer.spreadsheet;

import org.json.JSONArray;
import org.json.JSONObject;

import com.openxcell.util.StringUtils;

/**
 * @author vicky.thakor
 * @since 2018-05-15
 * 
 * @change allow wrap text on header column's cell
 * @author vicky.thakor
 * @since 2018-06-01
 * 
 * @change header comment
 * @author vicky.thakor
 * @since 2018-06-04
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
	private JSONArray wrapTextHeaders;
	private JSONObject headerComment;

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
			wrapTextHeaders = jsonObject.optJSONArray("wrap_text_headers");
			headerComment = jsonObject.optJSONObject("header_comments");
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
	
	public JSONArray getWrapTextHeaders() {
		return wrapTextHeaders;
	}
	
	public JSONObject getHeaderComment() {
		return headerComment;
	}
}
