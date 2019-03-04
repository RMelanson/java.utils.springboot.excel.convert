package utils.conversion.excel_to_json;

import java.io.IOException;

import utils.json.JacksonMarshaller;

public class Start {
	static String defaultDirectory = "C:\\Dev\\microservices\\data\\";
	static String defaultFileName = "Default";
	static String defaultInputType = ".xls";
	static String defaultOutputType = ".json";
	static String inputFile = defaultDirectory+defaultFileName+defaultInputType;
	static String outputFile = defaultDirectory+defaultFileName+defaultOutputType;
	
	public static void main(String[] args) throws IOException {
//		processArgs(args);
		String jsonString = Parse.xlsToJSON(inputFile);
		Parse.writeOutput(outputFile, jsonString);
	}

	private static void processArgs(String[] parms) {
		int parmsCount = parms.length;
		if (parmsCount == 1) {
			if (JacksonMarshaller.isValidJSON(parms[0])) {
			}	
		}
	}
}
