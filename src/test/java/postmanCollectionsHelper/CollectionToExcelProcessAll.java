package postmanCollectionsHelper;


import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class CollectionToExcelProcessAll {
	private static final int MAX_RESPONSE_LENGTH = 32000;
	private static Properties envProperties = new Properties();

	public static void main(String[] args) throws IOException {
		String inputFolder = "/Users/bs00880/myworkspace/collections";
		String outputExcel = "/Users/bs00880/myworkspace/output/API_Details_"+getClassNameStatic()+""+getDateTime()+".xlsx";
		String envFile = "environment.properties";

		envProperties = loadProperties(envFile);
		processPostmanCollections(inputFolder, outputExcel);
	}

	public static String getDateTime() {
		return LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
	}

	public static String getClassNameStatic() {
		return CollectionToExcelProcessAll.class.getSimpleName();  // Returns "MyClass"
	}


	private static Properties loadProperties(String filePath) {
		Properties properties = new Properties();
		try (InputStream input = new FileInputStream(filePath)) {
			properties.load(input);
		} catch (FileNotFoundException e) {
			System.out.println("Environment properties file not found. Using default values.");
		} catch (IOException e) {
			System.out.println("Error loading environment properties file: " + e.getMessage());
		}
		return properties;
	}

	private static void processPostmanCollections(String inputFolder, String outputExcel) throws IOException {
		File folder = new File(inputFolder);
		File[] files = folder.listFiles((dir, name) -> name.endsWith(".json"));
		if (files == null || files.length == 0) {
			System.out.println("No Postman collections found in the folder.");
			return;
		}

		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("API Details");
		Row headerRow = sheet.createRow(0);
		String[] headers = {"File Name","Full Directory Path", "Parent Folder", "Feature Name","Request Method", "Request Body", "cURL",
				"Response Count", "Response body", "Response headers", "status", "Status code","Response-Status,Body,Headers"};

		for (int i = 0; i < headers.length; i++) {
			headerRow.createCell(i).setCellValue(headers[i]);
		}


		int rowNum = 1;
		ObjectMapper objectMapper = new ObjectMapper();

		for (File file : files) {
			JsonNode rootNode = objectMapper.readTree(file);
			JsonNode itemArray = rootNode.path("item");
			rowNum = processItems(itemArray, "", "", sheet, rowNum, file.getName(), new HashSet<>());
		}

		try (FileOutputStream fileOut = new FileOutputStream(outputExcel)) {
			workbook.write(fileOut);
		}
		workbook.close();
		System.out.println("API details exported to " + outputExcel);
	}

	private static int processItems(JsonNode items, String parentPath, String parentFolder, Sheet sheet, int rowNum,
			String fileName, Set<String> uniqueRequests) {
		Iterator<JsonNode> elements = items.elements();
		while (elements.hasNext()) {
			JsonNode item = elements.next();
			String itemName = item.path("name").asText();
			JsonNode request = item.path("request");

			if (request.isMissingNode()) {
				rowNum = processItems(item.path("item"), parentPath + itemName + "/", itemName, sheet, rowNum, fileName, uniqueRequests);
			} else {
				//				System.out.println(request.path("url").toString());
				String fullPath = parentPath + itemName;
				String featureName = extractFeatureName(request.path("url"));
				String extractedApiRequest = extractApiRequest(request.path("url"));
				String curlCommand = generateCurlCommand(request);
				String requestMethod = request.path("method").asText();
				String requestBody = extractRequestBody(request.path("url").toString(), request);
				String requestKey = requestMethod + " " + extractedApiRequest;

				// Extract saved responses with details (status, code, headers, body)
				List<Map<String, String>> savedResponses = extractSavedResponses(item);
				int responseCount = savedResponses.size();



				if (savedResponses.isEmpty()) {
					// Ensure at least one row even if no saved responses exist
					Map<String, String> emptyResponse = new HashMap<>();
					emptyResponse.put("body", "");
					emptyResponse.put("code", "");
					emptyResponse.put("status", "");
					emptyResponse.put("headers", "");
					savedResponses.add(emptyResponse);
				}
				// single row only
				//					 if (!savedResponses.isEmpty()) {
				//			                Map<String, String> responseDetails = savedResponses.get(0); // Get only the first response
				//			                
				//			                Row row = sheet.createRow(rowNum++);
				//
				//			                row.createCell(0).setCellValue(fileName);
				//			                row.createCell(1).setCellValue(fullPath);
				//			                row.createCell(2).setCellValue(parentFolder);
				//			                row.createCell(3).setCellValue(featureName);
				//			                row.createCell(4).setCellValue(requestMethod);
				//			                row.createCell(5).setCellValue(requestBody);
				//			                row.createCell(6).setCellValue(truncateResponse(request.path("url").toString(), curlCommand));
				//
				//			                // Extract response details
				//			                String responseBody = responseDetails.getOrDefault("body", "");
				//			                String responseHeaders = responseDetails.getOrDefault("headers", "");
				//			                String responseStatus = responseDetails.getOrDefault("status", "");
				//			                String responseCode = responseDetails.getOrDefault("code", "");
				//
				//			                // Truncate response if necessary
				//			                responseBody = truncateResponse(request.path("url").toString(), responseBody);
				//
				//			                // Summary column
				//			                String restCombo = responseCode + System.lineSeparator()
				//			                                 + responseHeaders + System.lineSeparator()
				//			                                 + responseBody;
				//			                restCombo = truncateResponse(request.path("url").toString(), restCombo);
				//
				//			                row.createCell(7).setCellValue(responseCount);  
				//			                row.createCell(8).setCellValue(responseBody);
				//			                row.createCell(9).setCellValue(responseHeaders);
				//			                row.createCell(10).setCellValue(responseStatus);
				//			                row.createCell(11).setCellValue(responseCode);
				//			                row.createCell(12).setCellValue(restCombo);
				//			            }
				// Loop through all saved responses instead of picking only the first one
				for (Map<String, String> responseDetails : savedResponses) {
					Row row = sheet.createRow(rowNum++);

					row.createCell(0).setCellValue(fileName);
					row.createCell(1).setCellValue(fullPath);
					row.createCell(2).setCellValue(parentFolder);
					row.createCell(3).setCellValue(featureName);
					row.createCell(4).setCellValue(requestMethod);
					row.createCell(5).setCellValue(requestBody);

					if (curlCommand != null && curlCommand.length() > MAX_RESPONSE_LENGTH) {
						List<String> partsOfCurlCommand  = splitString(curlCommand, MAX_RESPONSE_LENGTH);
						//split text into 13,14
						int resinexstart=13;
						for (int i = 0; i < Math.min(2, partsOfCurlCommand.size()); i++) {
							//				                System.out.println("Part " + (i + 1) + ": " + partsOfResponseBody.get(i));

							row.createCell(resinexstart+i).setCellValue(partsOfCurlCommand.get(i));
							System.out.println(resinexstart+i);
						}
					}

					row.createCell(6).setCellValue(truncateResponse(request.path("url").toString(), curlCommand));

					// Extract response details
					String responseBody = responseDetails.getOrDefault("body", "");
					String responseHeaders = responseDetails.getOrDefault("headers", "");
					String responseStatus = responseDetails.getOrDefault("status", "");
					String responseCode = responseDetails.getOrDefault("code", "");


					if (responseBody != null && responseBody.length() > MAX_RESPONSE_LENGTH) {
						List<String> partsOfResponseBody  = splitString(responseBody, MAX_RESPONSE_LENGTH);
						//split text into 15,16,17
						int responseBodyIndex=15;
						for (int i = 0; i < Math.min(3, partsOfResponseBody.size()); i++) {
							//				                System.out.println("Part " + (i + 1) + ": " + partsOfResponseBody.get(i));

							row.createCell(responseBodyIndex+i).setCellValue(partsOfResponseBody.get(i));

						}
					}
					
					String nameofSavedResponse = responseDetails.getOrDefault("name", "");
					
					String[] parts = nameofSavedResponse.split("-");
					row.createCell(21).setCellValue(parts.length > 0 ? parts[0] : "");
					row.createCell(22).setCellValue(parts.length > 1 ? parts[1] : "");
					
					
					// Truncate response if necessary
					responseBody = truncateResponse(request.path("url").toString(), responseBody);

					// Summary column
					String restCombo = responseCode + System.lineSeparator()
					+ responseHeaders + System.lineSeparator()
					+ responseBody;
					restCombo = truncateResponse(request.path("url").toString(), restCombo);

					if (restCombo != null && restCombo.length() > MAX_RESPONSE_LENGTH) {
						List<String> partsOfRestCombo  = splitString(responseBody, MAX_RESPONSE_LENGTH);
						//split text into 18,19,20
						int restComboIndex=18;
						for (int i = 0; i < Math.min(3, partsOfRestCombo.size()); i++) {
							//				                System.out.println("Part " + (i + 1) + ": " + partsOfResponseBody.get(i));

							row.createCell(restComboIndex+i).setCellValue(partsOfRestCombo.get(i));

						}
					}

					row.createCell(7).setCellValue(responseCount);
					row.createCell(8).setCellValue(responseBody);
					row.createCell(9).setCellValue(responseHeaders);
					row.createCell(10).setCellValue(responseStatus);
					row.createCell(11).setCellValue(responseCode);
					row.createCell(12).setCellValue(restCombo);
				}

			}
		}
		return rowNum;
	}

	// Method to split the large string into chunks of a given size
	public static List<String> splitString(String input, int chunkSize) {
		List<String> parts = new ArrayList<>();
		int length = input.length();

		for (int i = 0; i < length; i += chunkSize) {
			parts.add(input.substring(i, Math.min(length, i + chunkSize)));
		}

		return parts;
	}

	// Method to print only the first three parts if available
	public static void printFirstThree(List<String> parts) {
		for (int i = 0; i < Math.min(3, parts.size()); i++) {
			System.out.println("Part " + (i + 1) + ": " + parts.get(i));
		}
	}



	private static String generateCurlCommand(JsonNode requestNode) {
		String method = requestNode.path("method").asText();
		JsonNode urlNode = requestNode.path("url");
		String url = urlNode.path("raw").asText();

		// Replace environment variables in the URL
		for (String key : envProperties.stringPropertyNames()) {
			url = url.replace("{{" + key + "}}", envProperties.getProperty(key));
		}

		StringBuilder curl = new StringBuilder("curl -X ").append(method).append(" \"").append(url).append("\"");

		// Process headers
		JsonNode headers = requestNode.path("header");
		boolean hasJsonHeader = false;
		for (JsonNode header : headers) {
			String headerKey = header.path("key").asText();
			String headerValue = header.path("value").asText();

			if (!headerKey.isEmpty() && !headerValue.isEmpty()) {
				curl.append(" -H \"").append(headerKey).append(": ").append(headerValue).append("\"");
			}

			if (headerKey.equalsIgnoreCase("Content-Type") && headerValue.equalsIgnoreCase("application/json")) {
				hasJsonHeader = true;
			}
		}

		// Process body for POST, PUT, PATCH methods
		JsonNode body = requestNode.path("body");
		boolean isPostPutPatch = method.equalsIgnoreCase("POST") || method.equalsIgnoreCase("PUT") || method.equalsIgnoreCase("PATCH");

		if (isPostPutPatch && !body.isMissingNode() && "raw".equals(body.path("mode").asText())) {
			// Check if "options" exist to determine JSON type
			boolean isJsonBody = !body.has("options") || "json".equals(body.path("options").path("raw").path("language").asText());

			String rawBody = body.path("raw").asText();
			if (!rawBody.isEmpty() && isJsonBody) {
				if (!hasJsonHeader) {
					curl.append(" -H \"Content-Type: application/json\"");
				}
				curl.append(" --data '").append(escapeSingleQuotes(formatJson(rawBody))).append("'");
			}
		}

		return curl.toString();
	}

	/**
	 * Escapes single quotes in JSON to prevent cURL errors.
	 */
	private static String escapeSingleQuotes(String json) {
		return json.replace("'", "'\"'\"'"); // Corrected escaping for cURL
	}

	/**
	 * Formats JSON string (optional, can be removed if not needed).
	 */
	private static String formatJson(String json) {
		try {
			ObjectMapper mapper = new ObjectMapper();
			Object jsonObj = mapper.readValue(json, Object.class);
			return mapper.writeValueAsString(jsonObj);
		} catch (Exception e) {
			return json; // Return original if formatting fails
		}
	}


	private static String extractFeatureName(JsonNode urlNode) {
		String rawUrl = urlNode.path("raw").asText();
		int queryIndex = rawUrl.indexOf("?");
		String featureUrl = (queryIndex > 0) ? rawUrl.substring(0, queryIndex) : rawUrl;
		featureUrl = featureUrl.replaceAll("\\{\\{base_url\\}}|\\{\\{baseUrl\\}}|https://mygp-dev\\.|https://mygptest\\.grameenphone|\\{\\{catalog_ms_url\\}}", "");
		return featureUrl;
	}

	private static String extractApiRequest(JsonNode urlNode) {
		String rawUrl = urlNode.path("raw").asText(); // Extract full raw URL

		int queryIndex = rawUrl.indexOf("?");

		// If there's a query parameter, consider everything before it
		return (queryIndex > 0) ? rawUrl.substring(0, queryIndex) : rawUrl;
	}


	private static String extractRequestBody(String url, JsonNode requestNode) {
		JsonNode body = requestNode.path("body");

		try {
			if (!body.isMissingNode() && body.has("mode") && "raw".equals(body.path("mode").asText())) {
				String rawBody = body.path("raw").asText();

				// If "options" exists, check for language = json
				if (body.has("options") && "json".equals(body.path("options").path("raw").path("language").asText())) {
					return rawBody.isEmpty() ? "" : truncateResponse(url,formatJson(rawBody));
				}

				// If "options" is missing, still extract the raw body
				return rawBody.isEmpty() ? "" : truncateResponse(url,formatJson(rawBody));
			}
		} catch (Exception e) {
			System.err.println("Error extracting request body: " + e.getMessage());
		}

		return "";
	}



	//    private static List<String> extractSavedResponses(JsonNode itemNode) {
	//        List<String> responses = new ArrayList<>();
	//        JsonNode responseArray = itemNode.path("response");
	//
	//        if (responseArray.isArray()) {
	//            for (JsonNode responseNode : responseArray) {
	//                JsonNode bodyNode = responseNode.path("body");
	//                if (!bodyNode.isMissingNode()) {
	//                    responses.add(bodyNode.asText());
	//                }
	//            }
	//        }
	//        return responses;
	//    }

	private static List<Map<String, String>> extractSavedResponses(JsonNode itemNode) {
		List<Map<String, String>> responses = new ArrayList<>();
		JsonNode responseArray = itemNode.path("response");

		if (responseArray.isArray()) {
			for (JsonNode responseNode : responseArray) {
				Map<String, String> responseDetails = new HashMap<>();

				// Extract response body
				JsonNode bodyNode = responseNode.path("body");
				responseDetails.put("body", bodyNode.isMissingNode() ? "" : truncateResponse("saved response",bodyNode.asText()));

				// Extract response code
				responseDetails.put("code", responseNode.path("code").asText());

				// Extract response status
				responseDetails.put("status", responseNode.path("status").asText());
				responseDetails.put("name", responseNode.path("name").asText());

				// Extract response headers
				StringBuilder headersBuilder = new StringBuilder();
				JsonNode headersArray = responseNode.path("header");
				if (headersArray.isArray()) {
					for (JsonNode header : headersArray) {
						headersBuilder.append(header.path("key").asText())
						.append(": ")
						.append(header.path("value").asText())
						.append("; ");
					}
				}
				responseDetails.put("headers", headersBuilder.toString());

				responses.add(responseDetails);
			}
		}
		return responses;
	}


	private static String truncateResponse(String url, String response) {

		if (response != null && response.length() > MAX_RESPONSE_LENGTH) {
			System.out.println(url+" - - - "+response.length());
			return response.substring(0, MAX_RESPONSE_LENGTH) + "... [ truncated ]";
		}
		return response;
	}
}


