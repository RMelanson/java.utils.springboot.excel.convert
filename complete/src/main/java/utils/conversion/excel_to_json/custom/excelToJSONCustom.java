package utils.conversion.excel_to_json.custom;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

import utils.json.JacksonMarshaller;

class Parse {

	static String xlsToJSON(String inputFile) {
		String jsonString = null;
		File iFile = new File(inputFile);
		LinkedHashMap<String, ArrayList<Object>> xlsFileReqMap = new LinkedHashMap<String, ArrayList<Object>>();
		HSSFWorkbook workbook = null;
		try {
			workbook = new HSSFWorkbook(new FileInputStream(iFile));
			xlsFileReqMap = getSheetMaps(workbook);
			buildJsonXlsMap(iFile, xlsFileReqMap);
			jsonString = JacksonMarshaller.mapJsonString(xlsFileReqMap);
			System.out.println(jsonString);
		} catch (FileNotFoundException e) {
			System.err.println("FileNotFoundException" + e.getMessage());
		} catch (IOException e) {
			System.err.println("IOException" + e.getMessage());
		} catch (Exception e) {
			System.err.println("Exception" + e.getMessage());
		} finally {
			try {
				if (workbook != null) {
					workbook.close();
				}
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return jsonString;
	}

	static void writeOutput(String outputFile, String jsonString) {
		try {
			File oFile = new File(outputFile);
			FileOutputStream fos = new FileOutputStream(oFile);
			fos.write(jsonString.getBytes());
			fos.close();
		} catch (FileNotFoundException e) {
			System.err.println("Exception" + e.getMessage());
		} catch (IOException e) {
			System.err.println("Exception" + e.getMessage());
		} finally {
		}
	}

	private static void buildJsonXlsMap(File iFile, LinkedHashMap<String, ArrayList<Object>> xlsFileReqMap) {
		HSSFWorkbook workbook = null;
		try {
			workbook = new HSSFWorkbook(new FileInputStream(iFile));

			for (Entry<String, ArrayList<Object>> sheet : xlsFileReqMap.entrySet()) {
				getSheetArr(workbook, sheet);
			}

		} catch (FileNotFoundException e) {
			System.err.println("Exception" + e.getMessage());
		} catch (IOException e) {
			System.err.println("Exception" + e.getMessage());
		} finally {
			try {
				workbook.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	private static void getSheetArr(HSSFWorkbook workbook, Entry<String, ArrayList<Object>> sheetEntry) {

		String sheetName = sheetEntry.getKey();
//		System.out.println("Sheet Name = " + sheetName);
		HSSFSheet sheet = workbook.getSheet(sheetName);

		int rowNum = sheet.getFirstRowNum();
		Row row = sheet.getRow(rowNum);
		ArrayList<Object> headerKeys = (ArrayList<Object>) sheetEntry.getValue().clone();
		ArrayList<Object> sheetArr = getSheetArr(sheet, headerKeys);
		sheetEntry.setValue(sheetArr);
	}

	private static ArrayList<Object> getSheetArr(HSSFSheet sheet, ArrayList<Object> headerKeys) {
		ArrayList<Object> sheetArr = new ArrayList<Object>();
		int rowNum = sheet.getFirstRowNum();
		Row row = sheet.getRow(rowNum);
		while (++rowNum <= sheet.getLastRowNum()) {
			row = sheet.getRow(rowNum);
			LinkedHashMap<Object, Object> dataRowMap = getDataRowMap(headerKeys, row);
			sheetArr.add(rowNum - 1, dataRowMap);
		}
		return sheetArr;
	}

	private static LinkedHashMap<String, ArrayList<Object>> getSheetMaps(HSSFWorkbook workbook) {
		LinkedHashMap<String, ArrayList<Object>> sheetMaps = new LinkedHashMap<String, ArrayList<Object>>();
		int numSheets = workbook.getNumberOfSheets();

		for (int cnt = 0; cnt < numSheets; cnt++) {
			HSSFSheet sheet = workbook.getSheetAt(cnt);
			String sheetName = sheet.getSheetName();
//			System.out.println("Sheet Name = " + sheetName);
			ArrayList<Object> sheetArr = getHeaderKeyArr(sheet);
			sheetMaps.put(sheetName, sheetArr);
		}
		return sheetMaps;
	}

	private static ArrayList<Object> getHeaderKeyArr(HSSFSheet sheet) {
		ArrayList<Object> headerArr = new ArrayList<Object>();
		int rowNum = sheet.getFirstRowNum();
		Row row = sheet.getRow(rowNum);
		row.getRowNum();
		for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
			Cell cell = row.getCell(cellNum);
			Object cellObjVal = null;
			try {
				switch (cell.getCellType()) {

				case BOOLEAN:
					cellObjVal = cell.getBooleanCellValue();
					break;

				case NUMERIC:
					if (DateUtil.isCellDateFormatted(cell)) {
						SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
						cellObjVal = dateFormat.format(cell.getDateCellValue());
					} else {
						cellObjVal = (Double) cell.getNumericCellValue();
						Double dblVal = (Double) cellObjVal;
						cellObjVal = dblVal.longValue();
					}
					break;

				case STRING:
					cellObjVal = cell.getRichStringCellValue().getString();
					break;

				case BLANK:
					cellObjVal = "NULL";
					break;
				default:
					break;
				}
				headerArr.add(cellNum, cellObjVal);
			} catch (NullPointerException e) {
				// do something clever with the exception
//				System.out.println("nullException" + e.getMessage());
			}
		}
		return headerArr;
	}

	private static LinkedHashMap<Object, Object> getDataRowMap(ArrayList<Object> headerKeys, Row row) {
		LinkedHashMap<Object, Object> dataMap = new LinkedHashMap<Object, Object>();
		LinkedHashMap<Object, Object> argsMap = null;
		row.getRowNum();
		for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
			Cell cell = row.getCell(cellNum);
			Object cellObjVal = "null";
			if (cell != null)
				try {
					switch (cell.getCellType()) {

					case BOOLEAN:
						cellObjVal = cell.getBooleanCellValue();
						break;

					case NUMERIC:
						if (DateUtil.isCellDateFormatted(cell)) {
							SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
							cellObjVal = dateFormat.format(cell.getDateCellValue());
						} else {
							cellObjVal = (Double) cell.getNumericCellValue();
							Double dblVal = (Double) cellObjVal;
							cellObjVal = dblVal.longValue();
						}
						break;

					case STRING:
						cellObjVal = cell.getRichStringCellValue().getString();
						break;

					case BLANK:
						cellObjVal = "NULL";
						break;
					default:
						break;
					}
				} catch (NullPointerException e) {
					// do something clever with the exception
					System.out.println("nullException" + e.getMessage());
				}
			if (cellNum != 1) {
				dataMap.put(headerKeys.get(cellNum), cellObjVal);
			} else {  // Create Parameter Array List
				ArrayList<Object> argsArr = new ArrayList<Object>();
				argsMap = new LinkedHashMap<Object, Object>();
				dataMap.put("args", argsArr);
				argsMap.put(headerKeys.get(cellNum), cellObjVal);
			}
		}
		return dataMap;
	}
}