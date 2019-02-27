package utils.conversion.excel_to_json;

import java.io.File;
import java.io.IOException;

public class main {
	static String defaultDirectory = "C:\\Dev\\microservices\\data\\";
	static String defaultFileName = "Default";
	static String defaultInputType = ".xls";
	static String defaultOutputType = ".json";
	static String defaultInputFile = defaultDirectory+defaultFileName+defaultInputType;
	static String defaultOutputFile = defaultDirectory+defaultFileName+defaultOutputType;
	
	public static void main(String[] args) throws IOException {
		File inputFile = new File(args.length > 1 ? args[1]: defaultInputFile);
		File outputFile = new File(args.length > 2 ? args[2]: defaultOutputFile);
		convert.xlsToJSON(inputFile, outputFile);
	}

}
