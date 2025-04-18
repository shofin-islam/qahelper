package postmanCollectionsHelper;


import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class CollectionToExcelProcessor {
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
        return CollectionToExcelProcessor.class.getSimpleName();  // Returns "MyClass"
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
        String[] headers = {"Full Directory Path", "Parent Folder", "Feature Name", "cURL", "File Name", "Saved Response",
                "Response Count", "Request Method", "JSON Request Body"};
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
            System.out.println("item Name: "+itemName);
            JsonNode request = item.path("request");

            if (request.isMissingNode()) {
                rowNum = processItems(item.path("item"), parentPath + itemName + "/", itemName, sheet, rowNum, fileName, uniqueRequests);
            } else {
            	System.out.println(request.path("url").toString());
                String fullPath = parentPath + itemName;
                String featureName = extractFeatureName(request.path("url"));
                String extractedApiRequest = extractApiRequest(request.path("url"));
                String curlCommand = generateCurlCommand(request);
                String requestMethod = request.path("method").asText();
                String requestBody = extractRequestBody(request);
                String requestKey = requestMethod + " " + extractedApiRequest;
                List<String> savedResponses = extractSavedResponses(item);
                int responseCount = savedResponses.size();

                if (!uniqueRequests.contains(requestKey)) {
                    uniqueRequests.add(requestKey);
                    if (savedResponses.isEmpty()) {
                        savedResponses.add(""); // Ensure at least one row
                    }

                    for (String response : savedResponses) {
                        Row row = sheet.createRow(rowNum++);
                        row.createCell(0).setCellValue(fullPath);
                        row.createCell(1).setCellValue(parentFolder);
                        row.createCell(2).setCellValue(featureName);
                        row.createCell(3).setCellValue(curlCommand);
                        row.createCell(4).setCellValue(fileName);
                        row.createCell(5).setCellValue(truncateResponse(response));
                        row.createCell(6).setCellValue(responseCount);
                        row.createCell(7).setCellValue(requestMethod);
                        row.createCell(8).setCellValue(requestBody);
                    }
                }
            }
        }
        return rowNum;
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

    
    private static String extractRequestBody(JsonNode requestNode) {
        JsonNode body = requestNode.path("body");
        
        if (!body.isMissingNode() && body.has("mode") && "raw".equals(body.path("mode").asText())) {
            // If "options" exists, check for language = json
            if (body.has("options") && "json".equals(body.path("options").path("raw").path("language").asText())) {
                String rawBody = body.path("raw").asText();
                return rawBody.isEmpty() ? "" : formatJson(rawBody);
            }
            
            // If "options" is missing, still extract the raw body
            String rawBody = body.path("raw").asText();
            return rawBody.isEmpty() ? "" : formatJson(rawBody);
        }

        return "";
    }

    
    private static List<String> extractSavedResponses(JsonNode itemNode) {
        List<String> responses = new ArrayList<>();
        JsonNode responseArray = itemNode.path("response");

        if (responseArray.isArray()) {
            for (JsonNode responseNode : responseArray) {
                JsonNode bodyNode = responseNode.path("body");
                if (!bodyNode.isMissingNode()) {
                    responses.add(bodyNode.asText());
                }
            }
        }
        return responses;
    }
    
    private static String truncateResponse(String response) {
        if (response != null && response.length() > MAX_RESPONSE_LENGTH) {
            return response.substring(0, MAX_RESPONSE_LENGTH) + "... [truncated]";
        }
        return response;
    }
}


