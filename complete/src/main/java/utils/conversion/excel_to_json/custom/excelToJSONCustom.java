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

	public static String xlsToJSON_Str(String inputXLSFile, String sheetName) {
		String jsonString = null;
		ArrayList<Object> xlsFileReqArr = xlsToJSON_Obj(inputXLSFile, sheetName);
		jsonString = JacksonMarshaller.mapJsonString(xlsFileReqArr);
		return jsonString;
	}

	static String xlsToJSON_Str(String inputXLSFile) {
		String jsonString = null;
		LinkedHashMap<String, ArrayList<Object>> xlsFileReqMap = xlsToJSON_Obj(inputXLSFile);
		jsonString = JacksonMarshaller.mapJsonString(xlsFileReqMap);
		return jsonString;
	}

	public static ArrayList<Object> xlsToJSON_Obj(String inputXLSFile, String sheetName) {
		LinkedHashMap<String, ArrayList<Object>> completeExcel = xlsToJSON_Obj(inputXLSFile);
		ArrayList<Object> sheetArr = completeExcel.get(sheetName);
		return sheetArr;
	}

	static LinkedHashMap<String, ArrayList<Object>> xlsToJSON_Obj(String inputFile) {
		File iFile = new File(inputFile);
		LinkedHashMap<String, ArrayList<Object>> xlsFileReqMap = new LinkedHashMap<String, ArrayList<Object>>();
		HSSFWorkbook workbook = null;
		try {
			workbook = new HSSFWorkbook(new FileInputStream(iFile));
			xlsFileReqMap = getSheetMaps(workbook);
			buildJsonXlsMap(iFile, xlsFileReqMap);
		} catch (FileNotFoundException e) {
			System.err.println("FileNotFoundException" + e.getMessage());
			System.out.println("ERROR FileNotFoundException: " + e.getMessage());
		} catch (IOException e) {
			System.err.println("IOException" + e.getMessage());
			System.out.println("ERROR IOException: " + e.getMessage());
		} catch (Exception e) {
			System.err.println("Exception" + e.getMessage());
			System.out.println("ERROR Exception: " + e.getMessage());
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
		return xlsFileReqMap;
	}

	static void writeOutput(String outputFile, String jsonString) {
		if (outputFile == null)
			System.out.println(jsonString);
		else
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
		Row currRow = sheet.getRow(rowNum);
		LinkedHashMap<Object, Object> prevDataRowMap = null;
		while (++rowNum <= sheet.getLastRowNum()) {
			currRow = sheet.getRow(rowNum);
			LinkedHashMap<Object, Object> currDataRowMap = getDataRowMap(headerKeys, currRow);
			if (currDataRowMap.get("function") != null) {
				sheetArr.add(currDataRowMap);
				prevDataRowMap = currDataRowMap;
			} else { // add to previous row args
				ArrayList<Object> args = (ArrayList<Object>) prevDataRowMap.get("args");
				args.add(currDataRowMap);
			}
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
		Boolean newFunction = true;
		row.getRowNum();
		for (int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
			Cell cell = row.getCell(cellNum);
			Object cellObjVal = "null";
			if (cell != null) {
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
				} catch (NullPointerException e) { // Do something clever with the null exception
					System.out.println("nullException" + e.getMessage());
				} catch (Exception e) { // Do something clever with the exception
					System.out.println("nullException" + e.getMessage());
				}
			}
			cellObjVal = cellObjVal.toString();
			String key = (String) headerKeys.get(cellNum);
			if (key.toUpperCase().equals("RET") || key.toUpperCase().equals("FUNCTION")) {
				if (!cellObjVal.toString().toUpperCase().equals("NULL"))
					dataMap.put(key, cellObjVal);
				else
					newFunction = false;

			} else {
				if (newFunction && cellNum == 1) { // Create Parameter Array List
					ArrayList<Object> argsArr = new ArrayList<Object>();
					argsMap = new LinkedHashMap<Object, Object>();
					argsArr.add(argsMap);
					dataMap.put("args", argsArr);
				}
				if (!cellObjVal.toString().toUpperCase().equals("NULL"))
					if (newFunction)
						argsMap.put(key, cellObjVal);
					else
						dataMap.put(key, cellObjVal);

			}
		}
		return dataMap;
	}
}