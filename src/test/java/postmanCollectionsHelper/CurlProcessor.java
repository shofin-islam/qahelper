package postmanCollectionsHelper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import io.restassured.RestAssured;
import io.restassured.http.ContentType;
import io.restassured.response.Response;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.*;
import com.fasterxml.jackson.databind.ObjectMapper;

public class CurlProcessor {
    private static Properties properties = new Properties();
    private static String CONFIGFILE="config.properties";

    public static void main(String[] args) throws Exception {
        loadProperties(CONFIGFILE);
        
//        processExcel("input.xlsx", "output.xlsx");
        
        String inputNames = getSheetNamesFromFile("input.xlsx");
        setPropertyValue("sheet_names", inputNames);
        System.out.println("sheet names: "+inputNames);
    }

   

 // Method to load properties from the given file
    private static void loadProperties(String filePath) throws IOException {
        try (InputStream input = new FileInputStream(filePath)) {
            properties.load(input);
        }
        System.out.println("Loaded properties: " + properties);
    }

    // Method to get a value by key from the properties
    private static String getPropertyValue(String key) {
        return properties.getProperty(key);
    }

    // Method to set a value for a key dynamically during execution
    private static void setPropertyValue(String key, String value) {
        properties.setProperty(key, value);

        // Save properties after modification without escaping characters
        try {
            saveProperties(CONFIGFILE);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

 // Method to save properties back to the file in alphabetical order by key
    private static void saveProperties(String filePath) throws IOException {
        // Create a TreeMap to automatically sort the properties alphabetically by key
        Map<String, String> sortedProperties = new TreeMap<>();
        for (Map.Entry<Object, Object> entry : properties.entrySet()) {
            String key = (String) entry.getKey();
            String value = (String) entry.getValue();
            sortedProperties.put(key, value);
        }

        // Write the sorted properties back to the file
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filePath))) {
            for (Map.Entry<String, String> entry : sortedProperties.entrySet()) {
                writer.write(entry.getKey() + "=" + entry.getValue());
                writer.newLine();
            }
        }
        System.out.println("Properties saved in alphabetical order: " + sortedProperties);
    }
    private static String getSheetNamesFromFile(String inputFile) throws Exception {
        Workbook workbook = new XSSFWorkbook(new FileInputStream(inputFile));
        List<String> sheetNames = new ArrayList<>();

        // Loop through all sheets and add their names
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            sheetNames.add(sheet.getSheetName()); // Get actual sheet name
        }

        workbook.close();

        // Return sheet names as a comma-separated string
        return String.join(",", sheetNames);
    }

    private static void processExcel(String inputFile, String outputFile) throws Exception {
        Workbook workbook = new XSSFWorkbook(new FileInputStream(inputFile));
        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Results");

        Row header = outputSheet.createRow(0);
        header.createCell(0).setCellValue("Sheet Name");
        header.createCell(1).setCellValue("Request cURL");
        header.createCell(2).setCellValue("Response File Path");

        int rowNum = 1;
        String[] sheetNames = properties.getProperty("sheetNames").split(",");

        // Create the response directory once for the entire execution
        String responseDirPath = createResponseDirectory();

        for (String sheetName : sheetNames) {
            Sheet sheet = workbook.getSheet(sheetName.trim());
            if (sheet == null) continue;

            System.out.println("Processing sheet: " + sheetName);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                Cell cell = row.getCell(3); // Assuming cURL request is in column C
                if (cell == null || cell.getStringCellValue().trim().isEmpty()) {
                    System.out.println("Skipping empty cURL cell at row: " + i);
                    continue; // Skip if the cell is empty
                }
                
                String curlRequest = cell.getStringCellValue();
                System.out.println("Original cURL: " + truncateString(curlRequest));

                String processedCurl = replacePlaceholders(curlRequest);
                System.out.println("Processed cURL: " + truncateString(processedCurl));

                String response = validateAndExecuteCurl(processedCurl);
                System.out.println("Response: " + truncateString(response));

                // Save the response to a JSON file in the dynamic directory
                String responseFilePath = writeResponseToFile(response, sheetName, i, responseDirPath);

                Row outputRow = outputSheet.createRow(rowNum++);
                outputRow.createCell(0).setCellValue(sheetName);
                outputRow.createCell(1).setCellValue(processedCurl); // Full cURL in Excel
                outputRow.createCell(2).setCellValue(responseFilePath); // Path to the response JSON file
            }
        }
        workbook.close();
        try (FileOutputStream outputStream = new FileOutputStream(outputFile)) {
            outputWorkbook.write(outputStream);
        }
        outputWorkbook.close();
    }


    private static String truncateString(String str) {
        if (str.length() > 500) {
            return str.substring(0, 500) + "..."; // Truncate to 500 characters and add ellipsis
        }
        return str;
    }

    private static String replacePlaceholders(String curl) {
        for (String key : properties.stringPropertyNames()) {
            curl = curl.replace("{{" + key + "}}", properties.getProperty(key));
        }
        return curl;
    }

    private static String validateAndExecuteCurl(String curl) {
        if (!isValidCurl(curl)) {
            System.out.println("Invalid cURL request: " + curl);
            return "Invalid cURL Request";
        }
        return executeApiRequest(curl);
    }

    private static boolean isValidCurl(String curl) {
        return curl.startsWith("curl") && (curl.contains("--request") || curl.contains("--location") || curl.contains("GET"));
    }

    private static String executeApiRequest(String curl) {
        try {
            Pattern urlPattern = Pattern.compile("'https?://[^\\s']+'");
            Pattern dataPattern = Pattern.compile("--data '(.+?)'", Pattern.DOTALL);
            Pattern methodPattern = Pattern.compile("--request (\\w+)");
            Pattern headerPattern = Pattern.compile("--header '([^']+): ([^']+)'", Pattern.DOTALL);

            Matcher urlMatcher = urlPattern.matcher(curl);
            Matcher dataMatcher = dataPattern.matcher(curl);
            Matcher methodMatcher = methodPattern.matcher(curl);
            Matcher headerMatcher = headerPattern.matcher(curl);

            String method = "GET"; // Default to GET
            if (methodMatcher.find()) {
                method = methodMatcher.group(1).toUpperCase();
            }

            if (urlMatcher.find()) {
                String url = urlMatcher.group().replace("'", "");
                String body = dataMatcher.find() ? dataMatcher.group(1) : "";

                System.out.println("Executing API request to: " + url);
                System.out.println("Request Method: " + method);
                System.out.println("Request Body: " + body);

                io.restassured.specification.RequestSpecification request = RestAssured.given();
                request.contentType(ContentType.JSON);

                while (headerMatcher.find()) {
                    String headerName = headerMatcher.group(1);
                    String headerValue = headerMatcher.group(2);
                    request.header(headerName, headerValue);
                    System.out.println("Adding Header: " + headerName + " = " + headerValue);
                }

                Response response;
                if (method.equals("POST")) {
                    response = request.body(body).post(url);
                } else if (method.equals("PUT")) {
                    response = request.body(body).put(url);
                } else if (method.equals("DELETE")) {
                    response = request.delete(url);
                } else if (method.equals("PATCH")) {
                    response = request.body(body).patch(url);
                } else { // Default to GET
                    response = request.get(url);
                }

                return response.getBody().asString();
            }
        } catch (Exception e) {
            System.out.println("Execution Error: " + e.getMessage());
            return "Execution Error: " + e.getMessage();
        }
        return "Execution Failed";
    }

    private static String createResponseDirectory() {
        // Generate a timestamp for the directory
        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        String responseDirPath = "responses_" + timestamp;  // Directory name with timestamp

        // Create the directory if it doesn't exist
        File responseDir = new File(responseDirPath);
        if (!responseDir.exists()) {
            responseDir.mkdir();
        }
        
        System.out.println("Response directory created: " + responseDirPath);
        return responseDirPath; // Return the directory path
    }

    private static String writeResponseToFile(String response, String sheetName, int rowIndex, String responseDirPath) {
        try {
            // Define the file path for the response JSON file
            String responseFileName = sheetName + "_response_" + rowIndex + ".json"; // Unique file for each response
            String responseFilePath = responseDirPath + "/" + responseFileName;

            // Write the response to the file
            try (FileWriter writer = new FileWriter(responseFilePath)) {
                writer.write(response);
            }

            System.out.println("Response saved to: " + responseFilePath);
            return responseFilePath; // Return the path to the file
        } catch (IOException e) {
            System.out.println("Error writing response to file: " + e.getMessage());
            return "Error";
        }
    }
}






