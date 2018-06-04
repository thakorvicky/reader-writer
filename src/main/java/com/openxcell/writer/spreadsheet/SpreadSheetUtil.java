package com.openxcell.writer.spreadsheet;

import java.util.Date;
import java.util.Objects;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;

import com.openxcell.util.StringUtils;

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
	 * @change header comment
	 * @author vicky.thakor
	 * @since 2018-06-04
	 * 
	 * @param row
	 * @param column
	 * @param value
	 */
	public static void writeCell(Workbook workbook, Row row, int column, Object value, Font font, DataFormat dataFormat,
			CellStyle cellStyle, boolean wrapText, String cellComment) {
		if (value == null)
			return;
		// change because of xlsx-stream reader library
		String strValue = Objects.isNull(value) ? "" : String.valueOf(value);
		/* Create cell with blank value */
		Cell cell = CellUtil.createCell(row, column, strValue);
		/* Determine the Format of value (i.e: Text, Number, Date, Float, etc...) */
		ExcelCellType valueFormat = getValueFormat(value);

		if (cellStyle != null) {
			/* Set cell fonts (i.e: Type, Color, Bold, etc...) */
			if (font != null) {
				cellStyle.setFont(font);
			}
			
			cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
			cellStyle.setWrapText(wrapText);
			cell.setCellStyle(cellStyle);
		}
		
		cellComment(workbook, row, cell, cellComment);

		switch (valueFormat) {
		case TEXT:
			cell.setCellValue(value.toString());
			break;
		case INTEGER:
			try {
				cell.setCellValue(((Number) value).intValue());
			} catch (Exception e) {
				Integer integer = Integer.valueOf((String) value);
				cell.setCellValue((integer).intValue());
			}
			CellUtil.setCellStyleProperty(cell, CellUtil.DATA_FORMAT, HSSFDataFormat.getBuiltinFormat("#,##0"));
			break;
		case FLOAT:
			try {
				cell.setCellValue(((Number) value).doubleValue());
			} catch (Exception e) {
				Float floatNumber = Float.valueOf((String) value);
				cell.setCellValue(floatNumber.doubleValue());
			}
			CellUtil.setCellStyleProperty(cell, CellUtil.DATA_FORMAT, HSSFDataFormat.getBuiltinFormat("#,##0.00"));
			break;
		case DATE:
			cell.setCellValue((Date) value);
			CellUtil.setCellStyleProperty(cell, CellUtil.DATA_FORMAT, HSSFDataFormat.getBuiltinFormat("m/d/yy"));
			break;
		case MONEY:
			cell.setCellValue(((Number) value).intValue());
			CellUtil.setCellStyleProperty(cell, CellUtil.DATA_FORMAT, dataFormat.getFormat("$#,##0.00;$#,##0.00"));
			break;
		case PERCENTAGE:
			cell.setCellValue(((Number) value).doubleValue());
			CellUtil.setCellStyleProperty(cell, CellUtil.DATA_FORMAT, HSSFDataFormat.getBuiltinFormat("0.00%"));
			break;
		default:
			cell.setCellValue(value.toString());
			break;
		}
	}

	/**
	 * @author vicky.thakor
	 * @since 2018-06-01
	 * 
	 * @param workbook
	 * @param row
	 * @param cell
	 * @param strComment
	 */
	private static void cellComment(Workbook workbook, Row row, Cell cell, String strComment) {
		if(StringUtils.nonNullNotEmpty(strComment)) {
			CreationHelper factory = workbook.getCreationHelper();
			Drawing drawing = cell.getSheet().createDrawingPatriarch();
			
			// When the comment box is visible, have it show in a 1x3 space
			ClientAnchor anchor = factory.createClientAnchor();
			anchor.setCol1(cell.getColumnIndex());
			anchor.setCol2(cell.getColumnIndex()+1);
			anchor.setRow1(row.getRowNum());
			anchor.setRow2(row.getRowNum()+5);
			
			// Create the comment and set the text+author
			Comment comment = drawing.createCellComment(anchor);
			RichTextString str = factory.createRichTextString(strComment);
			comment.setString(str);
			comment.setAuthor("Orderhive");
			
			// Assign the comment to the cell
			cell.setCellComment(comment);
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
