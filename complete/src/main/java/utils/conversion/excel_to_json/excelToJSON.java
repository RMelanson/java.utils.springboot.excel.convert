package utils.conversion.excel_to_json;

//public class excelToJSON {
//
//	public static void main(String[] args) {
//		// TODO Auto-generated method stub
//
//	}
//
//}
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

class convert {

	static void xlsToJSON(File inputFile, File outputFile) {
		StringBuffer cellDData = new StringBuffer();
		String cellDDataString = null;
		try {
			FileOutputStream fos = new FileOutputStream(outputFile);

			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(inputFile));
			HSSFSheet sheet = workbook.getSheetAt(0);
			Cell cell = null;
			Row row;
			int previousCell;
			int currentCell;
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				previousCell = -1;
				currentCell = 0;
				row = rowIterator.next();
				System.out.println("ROW:-->");
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					// System.out.println("true" +cellIterator.hasNext());
					cell = cellIterator.next();
					currentCell = cell.getColumnIndex();

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