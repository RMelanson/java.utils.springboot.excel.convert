package utils.conversion.excel_to_json;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedHashMap;

import utils.json.JacksonMarshaller;

public class Start {
	static String defaultDirectory = "C:\\Dev\\microservices\\data\\";
	static String defaultFileName = "Default";
	static String defaultInputType = ".xls";
	static String jsonType = ".json";
	static String defaultBuildMapName = "jsonRequestALL";
	static String inputFile = defaultDirectory + defaultFileName + defaultInputType;
	static String outputFile = defaultDirectory + defaultFileName + jsonType;
	static String buildMap = defaultDirectory + defaultBuildMapName + jsonType;

	public static void main(String[] args) throws IOException {
//		processArgs(args);
		processBuildFile(buildMap);
		String jsonString = Parse.xlsToJSON(inputFile);
		Parse.writeOutput(outputFile, jsonString);
	}

	private static void processArgs(String[] parms) {
		int parmsCount = parms.length;
		if (parmsCount == 1) {
			String buildFile = parms[0];
			processBuildFile(buildFile);
		}
	}

	private static void processBuildFile(String buildFile) {
		LinkedHashMap<?, ?> builder = readBuildFile(buildFile);

		String inputFile = (String) builder.get("INPUT_FILE");
		String outputFile = (String) builder.get("OUTPUT_FILE");
		LinkedHashMap<?, ?> sheets = (LinkedHashMap<?, ?>) builder.get("SHEETS");
		Object sheets2 = (LinkedHashMap<?, ?>) builder.get("SHEETS");
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
				return builderMap;
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
