package utils.conversion.excel_to_json;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

import utils.json.JacksonMarshaller;

class convert {
	static void xlsToJSON(File inputFile, File outputFile) {
		LinkedHashMap<String, ArrayList<?>> xlsFileMap = new LinkedHashMap<String, ArrayList<?>>();
		HSSFWorkbook workbook = null;
		try {
			FileOutputStream fos = new FileOutputStream(outputFile);

			workbook = new HSSFWorkbook(new FileInputStream(inputFile));
			int numSheets = workbook.getNumberOfSheets();

			for (int cnt = 0; cnt < numSheets; cnt++) {
				HSSFSheet sheet = workbook.getSheetAt(cnt);
				String SheetName = sheet.getSheetName();
				System.out.println("Sheet Name = " + SheetName);
				ArrayList sheetArr = getSheetArr(sheet, fos);
				xlsFileMap.put(SheetName, sheetArr);
			}
			String jsonString = JacksonMarshaller.mapJsonString(xlsFileMap);
			System.out.println(jsonString);
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

	static ArrayList<?> getSheetArr(HSSFSheet sheet, FileOutputStream fos) {
		ArrayList<Object> sheetArr = new ArrayList<Object>();
		try {
			StringBuffer cellDData = new StringBuffer();
			int rowNum = sheet.getFirstRowNum();
			Row row = sheet.getRow(rowNum);
			ArrayList<Object> headerKeys = getHeaderKeyArr(row);
			while (++rowNum <= sheet.getLastRowNum()) {
				row = sheet.getRow(rowNum);
				LinkedHashMap<Object, Object> dataRowMap = getDataRowMap(headerKeys, row);
				sheetArr.add(rowNum-1,dataRowMap);
			}
			fos.write(cellDData.toString().getBytes());
			fos.close();

		} catch (IOException e) {
			System.err.println("Exception" + e.getMessage());
		}
		
		ArrayList<LinkedHashMap<String, String>> lev3 = new ArrayList<LinkedHashMap<String, String>>();
		LinkedHashMap<String, String> lev4 = new LinkedHashMap<String, String>();

		return sheetArr;
	}

	private static ArrayList<Object> getHeaderKeyArr(Row row) {
		ArrayList<Object> headerArr = new ArrayList<Object>();
		Iterator<Cell> cellIterator = row.cellIterator();
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

	private static LinkedHashMap<Object,Object> getDataRowMap(ArrayList<Object> headerKeys, Row row) {
		LinkedHashMap<Object,Object> dataMap = new LinkedHashMap<Object,Object>();
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
				dataMap.put(headerKeys.get(cellNum), cellObjVal);
			} catch (NullPointerException e) {
				// do something clever with the exception
				System.out.println("nullException" + e.getMessage());
			}
		}
		return dataMap;
	}

	static void xlsToCVS(File inputFile, File outputFile) {
		StringBuffer cellDData = new StringBuffer();
		try {
			FileOutputStream fos = new FileOutputStream(outputFile);

			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(inputFile));
			HSSFSheet sheet = workbook.getSheetAt(0);
			Cell cell = null;
			Row row;
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				row = rowIterator.next();
				System.out.println("ROW:-->");
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					// System.out.println("true" +cellIterator.hasNext());
					cell = cellIterator.next();

					System.out.println("CELL:-->" + cell.toString());

					try {
						switch (cell.getCellType()) {

						case BOOLEAN:
							cellDData.append(cell.getBooleanCellValue() + ",");
							System.out.println("boo" + cell.getBooleanCellValue());
							break;

						case NUMERIC:
							if (DateUtil.isCellDateFormatted(cell)) {

								// System.out.println(cell.getDateCellValue());
								SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
								String strCellValue = dateFormat.format(cell.getDateCellValue());
								// System.out.println("date:"+strCellValue);
								cellDData.append(strCellValue + ",");
							} else {
								System.out.println(cell.getNumericCellValue());
								Double value = cell.getNumericCellValue();
								Long longValue = value.longValue();
								String strCellValue1 = new String(longValue.toString());
								// System.out.println("number:"+strCellValue1);
								cellDData.append(strCellValue1 + ",");
							}
							// cellDData.append(cell.getNumericCellValue() + ",");
							// String i=(new java.text.DecimalFormat("0").format(
							// cell.getNumericCellValue()+"," ));
							// System.out.println("number"+cell.getNumericCellValue());
							break;

						case STRING:
							String out = cell.getRichStringCellValue().getString();
							cellDData.append(cell.getRichStringCellValue().getString() + ",");
							// System.out.println("string"+cell.getStringCellValue());
							break;

						case BLANK:
							cellDData.append("" + "THIS IS BLANK");
							System.out.print("THIS IS BLANK");
							break;

						default:
							break;
						}
					} catch (NullPointerException e) {
						// do something clever with the exception
						System.out.println("nullException" + e.getMessage());
					}

				}
				workbook.close();
				int len = cellDData.length() - 1;
//      System.out.println("length:"+len);
//      System.out.println("length1:"+cellDData.length());
				cellDData.replace(cellDData.length() - 1, cellDData.length(), "");
				cellDData.append("\n");
			}
			// cellDData.append("\n");

//String out=cellDData.toString();
//System.out.println("res"+out);

//String o = out.substring(0, out.lastIndexOf(","));
//System.out.println("final"+o);
			fos.write(cellDData.toString().getBytes());
//fos.write(cellDDataString.getBytes());
			fos.close();

		} catch (FileNotFoundException e) {
			System.err.println("Exception" + e.getMessage());
		} catch (IOException e) {
			System.err.println("Exception" + e.getMessage());
		}
	}

}