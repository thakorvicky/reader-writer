package com.openxcell.writer.spreadsheet;

import java.util.Date;
import java.util.Objects;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;

/**
 * @author vicky.thakor
 * @since 2018-05-15
 */
public class SpreadSheetUtil {
	private static final String strABCD = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

	public enum ExcelCellType {
		TEXT, INTEGER, FLOAT, DATE, MONEY, PERCENTAGE
	}
	
	/**
	 * Get the cell identity by providing position.<br/>
	 * <br/>
	 * Example:<br/>
	 * 1. position = 1 => A<br/>
	 * 2. position = 26 => Z<br/>
	 * 3. position = 28 => AB<br/>
	 * 
	 * @author vicky.thakor
	 * @date 28th April, 2015
	 * @param position
	 * @return
	 */
	public static String generateCellName(int position) {
		String strPosition = "";
		while (position > 0) {
			int d = position % 26;
			strPosition = (d == 0 ? 'Z' : strABCD.charAt(d > 0 ? d - 1 : 0)) + strPosition;
			position = (position - 1) / 26;
		}
		return strPosition;
	}

	/**
	 * Create cell where you want result using formula.<br/>
	 * Example: SUM(E2:E10)<br/>
	 * 
	 * @author vicky.thakor
	 * @date 28th April, 2015
	 * @param row
	 * @param column
	 * @param formula
	 * @param font
	 * @param objHSSFDataFormat
	 */
	public static void writeFormulaCell(Workbook workbook, Row row, int column, String formula, Font font, short objHSSFDataFormat) {
		/* Create cell with blank value */
		Cell objHSSFCell = row.createCell(column);
		objHSSFCell.setCellType(CellType.FORMULA);
		objHSSFCell.setCellFormula(formula);
		if (objHSSFDataFormat != 0) {
			CellUtil.setCellStyleProperty(objHSSFCell, CellUtil.DATA_FORMAT, objHSSFDataFormat);
		}
	}
	
	/**
	 * Write value of cell.<br/>
	 * <br/>
	 * 
	 * @author vicky.thakor
	 * @date 9th April, 2015
	 * 
	 * @change wrap text
	 * @author vicky.thakor
	 * @since 2018-06-01
	 * 
	 * @param row
	 * @param column
	 * @param value
	 */
	public static void writeCell(Workbook workbook, Row row, int column, Object value, Font font, DataFormat dataFormat,
			CellStyle cellStyle, boolean wrapText) {
		if (value == null)
			return;
		// change because of xlsx-stream reader library
		String strValue = Objects.isNull(value) ? "" : String.valueOf(value);
		/* Create cell with blank value */
		Cell objHSSFCell = CellUtil.createCell(row, column, strValue);
		/* Determine the Format of value (i.e: Text, Number, Date, Float, etc...) */
		ExcelCellType valueFormat = getValueFormat(value);

		if (cellStyle != null) {
			/* Set cell fonts (i.e: Type, Color, Bold, etc...) */
			if (font != null) {
				cellStyle.setFont(font);
			}
			
			cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
			cellStyle.setWrapText(wrapText);
			objHSSFCell.setCellStyle(cellStyle);
		}

		switch (valueFormat) {
		case TEXT:
			objHSSFCell.setCellValue(value.toString());
			break;
		case INTEGER:
			try {
				objHSSFCell.setCellValue(((Number) value).intValue());
			} catch (Exception e) {
				Integer integer = Integer.valueOf((String) value);
				objHSSFCell.setCellValue((integer).intValue());
			}
			CellUtil.setCellStyleProperty(objHSSFCell, CellUtil.DATA_FORMAT, HSSFDataFormat.getBuiltinFormat("#,##0"));
			break;
		case FLOAT:
			try {
				objHSSFCell.setCellValue(((Number) value).doubleValue());
			} catch (Exception e) {
				Float floatNumber = Float.valueOf((String) value);
				objHSSFCell.setCellValue(floatNumber.doubleValue());
			}
			CellUtil.setCellStyleProperty(objHSSFCell, CellUtil.DATA_FORMAT, HSSFDataFormat.getBuiltinFormat("#,##0.00"));
			break;
		case DATE:
			objHSSFCell.setCellValue((Date) value);
			CellUtil.setCellStyleProperty(objHSSFCell, CellUtil.DATA_FORMAT, HSSFDataFormat.getBuiltinFormat("m/d/yy"));
			break;
		case MONEY:
			objHSSFCell.setCellValue(((Number) value).intValue());
			CellUtil.setCellStyleProperty(objHSSFCell, CellUtil.DATA_FORMAT, dataFormat.getFormat("$#,##0.00;$#,##0.00"));
			break;
		case PERCENTAGE:
			objHSSFCell.setCellValue(((Number) value).doubleValue());
			CellUtil.setCellStyleProperty(objHSSFCell, CellUtil.DATA_FORMAT, HSSFDataFormat.getBuiltinFormat("0.00%"));
			break;
		default:
			objHSSFCell.setCellValue(value.toString());
			break;
		}
	}

	/**
	 * To identify the format of value (i.e: Number, Float, Date, etc...).<br/>
	 * <br/>
	 * 
	 * @author vicky.thakor
	 * @date 9th April, 2015
	 * 
	 * @change use `format_cell_value` to change the type of cell
	 * @author vicky.thakor
	 * @since 2018-05-31
	 * 
	 * @param {@link
	 * 			Object}
	 * @return {@link ExcelCellType}
	 */
	private static ExcelCellType getValueFormat(Object value) {
		if (value instanceof Float || value instanceof Double) {
			return ExcelCellType.FLOAT;
		} else if (value instanceof Integer || value instanceof Long) {
			return ExcelCellType.INTEGER;
		} else if (value instanceof Date) {
			return ExcelCellType.DATE;
		} else {
//			if(Objects.nonNull(value) && RegExUtils.isInteger(String.valueOf(value))) {
//				return ExcelCellType.INTEGER;
//			}else if(Objects.nonNull(value) && RegExUtils.isFloatingNumber(String.valueOf(value))) {
//				return ExcelCellType.FLOAT;
//			}else {
				return ExcelCellType.TEXT;
//			}
		}
	}
}
