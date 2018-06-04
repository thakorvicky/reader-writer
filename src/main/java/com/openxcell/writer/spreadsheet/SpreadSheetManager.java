package com.openxcell.writer.spreadsheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.UUID;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.openxcell.io.FileHolder;
import com.openxcell.util.JSONUtils;

/**
 * @author vicky.thakor
 * @since 2018-05-15
 * 
 * @change wrap cell text
 * @author vicky.thakor
 * @since 2018-06-01
 */
public class SpreadSheetManager {
	private static Logger logger = Logger.getLogger(SpreadSheetManager.class.getName());
	
	private SXSSFWorkbook workbook; // http://poi.apache.org/spreadsheet/how-to.html#sxssf
	private SXSSFSheet sheet;
	private Row headerRow;
	private Row dataRow;
	private CellStyle cellStyle;
	private DataFormat dataFormat;

	/* streaming */
	private boolean enableStream = false;
	private int streamRowBuffer = -1;

	/* counter */
	private int workbookRowCount = 0;
	private int sheetCount = 0;
	private int sheetRowCount = 0;
	private int columnCount = 0;
	private int cellCount = 0;
	private int sheetChangeThreshold = 65000;

	/* configuration */
	private boolean freezeHeader = false;
	private boolean boldHeader = false;

	private Map<String, Integer> headers = new LinkedHashMap<>(0);
	private Map<Integer, String> headerIndex = new HashMap<>();
	private Map<String, Short> headerBackgroundColor;
	private Map<String, Short> headerTextColor;

	private SpreadSheetTemplate spreadSheetTemplate;
	private List<String> summationBeforeNewSheet = new ArrayList<>(0);
	private List<String> wrapTextHeaders = new ArrayList<>();

	/* Build new workbook */
	public void buildWorkbook(SpreadSheetTemplate spreadSheetTemplate) {
		if (enableStream) {
			workbook = new SXSSFWorkbook(streamRowBuffer);
			((SXSSFWorkbook) workbook).setCompressTempFiles(true);
		} else {
			workbook = new SXSSFWorkbook(-1);
		}
		
		this.spreadSheetTemplate = spreadSheetTemplate;
		dataFormat = workbook.createDataFormat();
		cellStyle = workbook.createCellStyle();
		
		if (Objects.nonNull(spreadSheetTemplate.getSummationHeaders())) {
			summationBeforeNewSheet = JSONUtils.JSONArrayToList(String.class, spreadSheetTemplate.getSummationHeaders());
		}
		
		if(Objects.nonNull(spreadSheetTemplate.getWrapTextHeaders())) {
			wrapTextHeaders = JSONUtils.JSONArrayToList(String.class, spreadSheetTemplate.getWrapTextHeaders());
		}
	}

	/**
	 * Enable stream in excel
	 * 
	 * @param enableStream
	 * @param streamRowBuffer
	 */
	public void enableStream(boolean enableStream, int streamRowBuffer) {
		this.enableStream = enableStream;
		this.streamRowBuffer = streamRowBuffer;
		if (enableStream && streamRowBuffer <= 0) {
			throw new RuntimeException("Provide streamRowBuffer greater than 0.");
		}
	}

	/* Prepare new sheet */
	private void prepareSheet() {
		if (sheetRowCount == 0 || sheetRowCount > sheetChangeThreshold) {
			sheetCount++;
			sheetRowCount = 0;
			columnCount = 0;

			
			sheet = workbook.createSheet("Sheet " + sheetCount);
			sheet.trackAllColumnsForAutoSizing();
			
			headerRow = createRow(sheetRowCount);

			if (freezeHeader) {
				sheet.createFreezePane(0, 1);
			}

			sheetRowCount++;
			workbookRowCount++;
		}
	}

	/**
	 * Set limit when to create new sheet.
	 * 
	 * @param rowCount
	 */
	public void setSheetChangeThreshold(int rowCount) {
		if (rowCount > 65000) {
			throw new RuntimeException("You can't write more than 65000 rows in one sheet");
		}
		sheetChangeThreshold = rowCount;
	}

	/**
	 * Auto-resize given column index.
	 * 
	 * @param columnIndex
	 */
//	private void autoResizeColumn(int columnIndex) {
//		sheet.trackColumnForAutoSizing(columnIndex);
//	}

	/**
	 * Add header to current sheet
	 * 
	 * @change comment on header
	 * @author vicky.thakor
	 * @since 2018-06-04
	 */
	public int addHeader(String header) {
		prepareSheet();

		if (headers.containsKey(header)) {
			columnCount = headers.get(header);
		} else {
			/* Will be used when creating new sheet */
			columnCount = headers.size();
			headers.put(header, columnCount);
			headerIndex.put(columnCount, header);
		}

		Font font = workbook.createFont();
		font.setFontName("Times New Roman");

		if (Objects.nonNull(headerTextColor) && headerTextColor.containsKey(header)) {
			font.setColor(headerTextColor.get(header));
		}

		if (boldHeader) {
			font.setBold(true);
		}
		sheet.trackColumnForAutoSizing(columnCount);
		
		String comment = null;
		if(Objects.nonNull(spreadSheetTemplate.getHeaderComment())) {
			comment = spreadSheetTemplate.getHeaderComment().optString(header);
		}
		
		SpreadSheetUtil.writeCell(workbook, headerRow, columnCount, header, font, dataFormat,
				workbook.createCellStyle(), false, comment);
		return columnCount++;
	}

	/**
	 * Freeze Header.
	 */
	public void freezeHeader() {
		freezeHeader = true;
	}

	/**
	 * Make header font bold.
	 */
	public void boldHeader() {
		boldHeader = true;
	}

	/**
	 * Auto-resize columns before moving to next sheet
	 */
	private void autoResizeHeader() {
		if (headers != null) {
			headers.values().stream().forEach(columnIndex -> {
				sheet.autoSizeColumn(columnIndex);
			});
		}
	}

	/**
	 * Set header text color.
	 * 
	 * @param headerTextColor
	 */
	public void setHeaderTextColor(Map<String, Short> headerTextColor) {
		this.headerTextColor = headerTextColor;
	}

	/**
	 * Set header background color.
	 * 
	 * @param headerBackgroundColor
	 */
	public void setHeaderBackgroundColor(Map<String, Short> headerBackgroundColor) {
		this.headerBackgroundColor = headerBackgroundColor;
	}

	/**
	 * Set header color
	 * 
	 * @param headerColor
	 */
	private void setHeaderColor() {
		if (Objects.nonNull(headerBackgroundColor)) {
			Row headerRow = sheet.getRow(0);
			if (Objects.nonNull(headerRow)) {
				headerBackgroundColor.entrySet().stream().filter(Objects::nonNull).forEach(entrySet -> {
					String headerName = entrySet.getKey();
					Short color = entrySet.getValue();
					if (headers.containsKey(headerName)) {
						int columnIndex = headers.get(headerName);
						CellStyle cellStyle = headerRow.getCell(columnIndex).getCellStyle();
						cellStyle.setFillForegroundColor(color);
						cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
						headerRow.getCell(columnIndex).setCellStyle(cellStyle);
					}
				});
			}
		}
	}

	/**
	 * Add new row in current sheet
	 * 
	 * @return
	 */
	public int newRow() {
		if (sheetRowCount >= sheetChangeThreshold) {
			if (summationBeforeNewSheet != null) {
				lastRow();
				summationBeforeNewSheet.stream().forEach(header -> {
					summation(header, 0);
				});
			}

//			autoResizeHeader();
			prepareSheet();
			if (headers != null) {
				headers.keySet().stream().forEach(header -> {
					addHeader(header);
				});
			}
		}

		dataRow = createRow(sheetRowCount);
		cellCount = 0;
		sheetRowCount++;
		workbookRowCount++;
		return sheetRowCount;
	}

	/**
	 * @since 2018-05-17
	 * @return
	 */
	public int copyPreviousRow() {
		Row row = sheet.getRow(sheetRowCount-1);
		newRow();
		
		for (int i = 0; i < row.getLastCellNum(); i++) {
			addValueCell(row.getCell(i));
		}
		return sheetRowCount;
	}
	
	/**
	 * For internal purpose.
	 */
	private void lastRow() {
		dataRow = createRow(sheetRowCount);
		cellCount = 0;
		sheetRowCount++;
		workbookRowCount++;
	}

	/**
	 * Add cell to current row.
	 */
	public void addValueCell(Object value) {
		addValueCell(cellCount, value);
		cellCount++;
	}

	/**
	 * @param index
	 * @param value
	 */
	public void addValueCell(int index, Object value) {
		SpreadSheetUtil.writeCell(workbook, dataRow, index, value, null, dataFormat, cellStyle, doWrapText(index), null);
	}
	
	/**
	 * Add cell for given header in current row.
	 */
	public void addValueCell(String header, Object value) {
		if (dataRow != null) {
			if (headers != null && headers.containsKey(header)) {
				SpreadSheetUtil.writeCell(workbook, dataRow, headers.get(header), value, null, dataFormat, cellStyle, wrapTextHeaders.contains(header), null);
			}
		} else {
			throw new RuntimeException("Initialize row");
		}
	}

	/**
	 * Perform summation on given property
	 */
	public void summation(String header, int lastNrows) {
		if (headers != null && headers.containsKey(header)) {
			int propertyPosition = headers.get(header);
			int fromPosition = lastNrows != 0 ? sheetRowCount - lastNrows : 2;

			/*
			 * property position start with `0` in program where it start with `1` in excel
			 */
			String getExcelColumnIdentity = SpreadSheetUtil.generateCellName(propertyPosition + 1);
			/* SUM(E2:E795) */
			String summationFormula = "";
			summationFormula = "SUM(" + getExcelColumnIdentity + fromPosition + ":" + getExcelColumnIdentity + ""
					+ (sheetRowCount - 1) + ")";

			SpreadSheetUtil.writeFormulaCell(workbook, dataRow, propertyPosition, summationFormula, null,
					HSSFDataFormat.getBuiltinFormat("#,##0.00"));
		}
	}

	/**
	 * Perform required steps to close the workbook for final output.
	 */
	private void doClose() {
		if (summationBeforeNewSheet != null) {
			lastRow();
			summationBeforeNewSheet.stream().forEach(header -> {
				summation(header, 0);
			});
		}
		setHeaderColor();
		autoResizeHeader();
	}

	/**
	 * Close workbook and get file created at temporary location.
	 * 
	 * @param path
	 *            - provide directory name or will be create at temporary location
	 *            of OS. (Optional)
	 * @param filename
	 *            - provide filename along with extension or will be generate using
	 *            {@link UUID#randomUUID()}. (Optional)
	 * @param deleteOnExit
	 *            - whether you want to delete file or not when program exit from
	 *            JVM (Terminate).
	 * @return
	 * @throws IOException
	 */
	public void closeWorkbook(FileHolder fileHolder) throws IOException {
		doClose();
		try (OutputStream objOutputStream = new FileOutputStream(fileHolder);) {
			workbook.write(objOutputStream);
		} catch (IOException ex) {
			logger.log(Level.SEVERE, ex.getMessage(), ex);
			throw ex;
		} finally {
			if (workbook != null) {
				((SXSSFWorkbook) workbook).dispose();
				workbook.close();
			}
		}
	}

	/**
	 * @param rownum
	 * @return
	 */
	private Row createRow(int rownum) {
		Row row = sheet.createRow(rownum);
		return row;
	}
	
	/**
	 * @param columnIndex
	 * @return
	 */
	private boolean doWrapText(int columnIndex) {
		if(headerIndex.containsKey(columnIndex)) {
			if(columnIndex == 2) {
				System.out.println("Wait here");
			}
			return wrapTextHeaders.contains(headerIndex.get(columnIndex));
		}
		return false;
	}
	
	public SpreadSheetTemplate getSpreadSheetTemplate() {
		return spreadSheetTemplate;
	}
	
	public int getWorkbookRowCount() {
		return workbookRowCount;
	}
}
