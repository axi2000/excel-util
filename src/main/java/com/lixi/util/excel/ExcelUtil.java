package com.lixi.util.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

public class ExcelUtil {


	/**
	 *
	 * @param input excel 文件的inputStream
	 * @param keyMapping
	 * @return List<Map<String, Object>>
	 * @throws IOException
     */
	public static List<Map<String, Object>> excelToList(InputStream input,
														Map<String, String> keyMapping) throws IOException{
		return excelToList(input, keyMapping, 0);
	}

	/**
	 * 将一个excel转换为List格式,默认获取第一个sheet
	 * 
	 * @param input
	 * @return List<Map<String, Object>>
	 * @throws IOException
	 */
	public static List<Map<String, Object>> excelToList(InputStream input,
			Map<String, String> keyMapping, int sheetIndex) throws IOException {
		XSSFWorkbook book = new XSSFWorkbook(input);
		XSSFSheet sheet = book.getSheetAt(sheetIndex);
		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		int rownum = sheet.getLastRowNum();
		// 第一行，标题行
		XSSFRow firstRow = sheet.getRow(0);
		int cellNum = firstRow.getLastCellNum();
		String[] keys = new String[cellNum];
		for (int i = 0; i < cellNum; i++) {
			XSSFCell cell = firstRow.getCell(i);
			if (cell != null) {
				keys[i] = cell.toString().trim();
				if (null != keyMapping) {
					String key = keyMapping.get(keys[i]);
					if (null != key) {
						keys[i] = key;
					}
				}
			} else {
				keys[i] = "";
			}
		}
		// 后面的行，数据行
		for (int i = 1; i < rownum + 1; i++) {
			XSSFRow row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			cellNum = row.getLastCellNum();
			Map<String, Object> record = new LinkedHashMap<String, Object>();
			for (int j = 0; j < cellNum; j++) {
				XSSFCell cell = row.getCell(j);
				if (cell != null && StringUtil.isNotEmpty(keys[j])) {
					record.put(keys[j], cell.toString());
				}
			}
			list.add(record);
		}
		return list;
	}

	public static List<Map<String, Object>> excelToList(InputStream input)
			throws IOException {
		return excelToList(input, null);
	}

	/**
	 * 将一个List<Map<String, Object>> 格式的数据转换为excel
	 * @param list
	 * @return
	 * @throws IOException
     */
	public static XSSFWorkbook listToExcel(List<Map<String, Object>> list)
			throws IOException {
		Map<String, Object> firstRecord = list.get(0);
		XSSFWorkbook book = new XSSFWorkbook();
		XSSFSheet sheet = book.createSheet();
		XSSFRow row = sheet.createRow(0);
		int i = 0;
		for (String key : firstRecord.keySet()) {
			XSSFCell cell = row.createCell(i++);
			cell.setCellValue(key);
		}

		for (i = 0; i < list.size(); i++) {
			Map<String, Object> r = list.get(i);
			int j = 0;
			row = sheet.createRow(i + 1);
			for (Object o : r.values()) {
				XSSFCell cell = row.createCell(j++);
				cell.setCellValue(o.toString());
			}
		}

		return book;
	}

	public static boolean isMergedRegion(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return true;
				}
			}
		}
		return false;
	}
	
	public static boolean isMergedRegion(Cell cell){
		
		return isMergedRegion(cell.getSheet(), cell.getRowIndex(), cell.getColumnIndex());
	}
}
