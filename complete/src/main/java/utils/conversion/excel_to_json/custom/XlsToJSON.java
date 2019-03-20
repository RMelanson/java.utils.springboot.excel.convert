package utils.conversion.excel_to_json.custom;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.LinkedHashMap;

import utils.json.JacksonMarshaller;

public class XlsToJSON {
	static String EOLN = "\n";
	static String defaultDirectory = "C:\\Dev\\microservices\\data\\";
	static String defaultFileName = "Default";
	static String defaultInputType = ".xls";
	static String jsonType = ".json";
	static String defaultBuildMapName = "jsonRequestALL";
	static String inputXLSFile = defaultDirectory + defaultFileName + defaultInputType;
	static String outputJSONFile = defaultDirectory + defaultFileName + jsonType;
	static String sheetName = null;
	static String buildMap = defaultDirectory + defaultBuildMapName + jsonType;

	public static void main(String[] args) throws IOException {
		processArgs(args);
		String jsonString;
		if (sheetName == null)
		   jsonString = Parse.xlsToJSON_Str(inputXLSFile);
		else
		   jsonString = Parse.xlsToJSON_Str(inputXLSFile,sheetName);
		Parse.writeOutput(outputJSONFile, jsonString);
	}

	private static void processArgs(String[] parms) {
		int parmsCount = parms.length;
		if ((parmsCount % 2) != 0 || parmsCount < 2) {
			usage();
			System.exit(-1);
		}
		for (int idx = 0; idx < parmsCount; idx += 2) {
			switch (parms[idx]) {
			case "-i":
				inputXLSFile = parms[idx + 1];
				break;
			case "-o":
				outputJSONFile = parms[idx + 1];
				break;
			case "-d":
				buildMap = parms[idx + 1];
				readBuildFile(buildMap);
				break;
			case "-s":
				sheetName = parms[idx + 1];
				break;
			default:
				usage();
				System.exit(-1);
			}
		}
	}

	private static void usage() {
		String usage = "XUsage:" + EOLN;
		usage += "XlsToJason <Parameters ( At least one of input or definition File Path Required)>" + EOLN;
		usage += "Parameters:\n" + EOLN;
		usage += "-i inputFilePathName (manditory if not defined in definition file)" + EOLN;
		usage += "-o outputFilePathName (optional)" + EOLN;
		usage += "-d definitionFilePath Name (optional JSON file defining input, and output files with required fields";
		System.out.println(usage);
	}

	private static void setBuildFileArgs(LinkedHashMap<?, ?> builderMap) {
//		LinkedHashMap<?, ?> builder = readBuildFile(buildFile);

		String inputFile = (String) builderMap.get("INPUT_FILE");
		String outputFile = (String) builderMap.get("OUTPUT_FILE");
//		LinkedHashMap<?, ?> sheets = (LinkedHashMap<?, ?>) builder.get("SHEETS");
//		Object sheets2 = (LinkedHashMap<?, ?>) builder.get("SHEETS");
	}

	private static LinkedHashMap<?, ?> readBuildFile(String buildFile) {
		LinkedHashMap<?, ?> builderMap = null;
		File bFile = new File(buildFile);
		int ch;
		try {
			StringBuffer strContent = new StringBuffer("");
			FileInputStream inputBuildFile = new FileInputStream(bFile);
			while ((ch = inputBuildFile.read()) != -1)
				strContent.append((char) ch);

			if (JacksonMarshaller.isValidJSON(strContent.toString())) {
				builderMap = (LinkedHashMap<?, ?>) JacksonMarshaller.jsonStringToClass(strContent.toString(),
						LinkedHashMap.class);
				setBuildFileArgs(builderMap);
			}
			inputBuildFile.close();
		} catch (FileNotFoundException e) {
			System.out.println("File " + bFile.getAbsolutePath() + " could not be found on filesystem");
		} catch (IOException ioe) {
			System.out.println("Exception while reading the file" + ioe);
		} finally {
		}
		return builderMap;
	}
}
