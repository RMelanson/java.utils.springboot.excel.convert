package utils.conversion.excel_to_json.json;

import java.io.IOException;

import org.apache.commons.lang3.StringUtils;
import org.codehaus.jackson.JsonGenerationException;
import org.codehaus.jackson.map.JsonMappingException;
import org.codehaus.jackson.map.ObjectMapper;
import org.codehaus.jackson.map.ObjectWriter;

public class JacksonMarshaller {

	public static String toJsonString(Object obj) {
		// mapJsonString(obj);

		ObjectWriter ow = new ObjectMapper().writer().withDefaultPrettyPrinter();
		String jsonString = null;
		try {
			jsonString = ow.writeValueAsString(obj);
		} catch (IOException e) {
			e.printStackTrace();
		}
		if (jsonString != null) {
			jsonString = jsonString.replaceAll("\\r", "");
		}
		return jsonString;
	}

	/*
	 * public static String toGsonString(Object obj) { Gson gson = new Gson();
	 * String gsonString = gson.toJson(obj); return gsonString ; }
	 */
	public static String mapJsonString(Object obj) {
		String jsonInString = null;
		try {
			// Convert object to JSON string
			ObjectMapper mapper = new ObjectMapper();
			jsonInString = mapper.writeValueAsString(obj);
			System.out.println(jsonInString);

			// Convert object to JSON string and pretty print
			jsonInString = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(obj);
			System.out.println(jsonInString);
		} catch (JsonGenerationException e) {
			e.printStackTrace();
		} catch (JsonMappingException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return jsonInString;
	}

	public static boolean isValidJSON(String JSON_STRING) {
		try {
			if (StringUtils.isEmpty(JSON_STRING))
				return false;
			final ObjectMapper mapper = new ObjectMapper();
			mapper.readTree(JSON_STRING);
			return true;
		} catch (IOException e) {
			return false;
		}
	}

	public static Object jsonStringToClass(String jsonString) {
		Object classObj = null;
		try {
			classObj = jsonStringToClass(jsonString, Object.class);
		} catch (Exception e) {
			e.printStackTrace();
			// throw e;
		}
		return classObj;
	}

	public static Object jsonStringToClass(String jsonString, Class<?> c) {
		Object classObj = null;
		if (!StringUtils.isEmpty(jsonString))
			try {
				classObj = new ObjectMapper().readValue(jsonString, c);
			} catch (Exception e) {
//				e.printStackTrace();
				return null;
			}
		return classObj;
	}
}
