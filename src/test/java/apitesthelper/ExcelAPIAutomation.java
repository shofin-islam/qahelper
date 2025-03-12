package apitesthelper;
import io.restassured.RestAssured;
import io.restassured.response.Response;
import io.restassured.specification.RequestSpecification;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelAPIAutomation {
    private static String authToken;
    private static String categoriesUrl;
    private static String excelFilePath = "api_request.xlsx";

    public static void main(String[] args) {
        System.out.println("Starting Excel API Automation...");
        loadProperties();
        processExcelData();
        System.out.println("Process completed.");
    }

    private static void loadProperties() {
        try (FileInputStream fis = new FileInputStream("config.properties")) {
            Properties properties = new Properties();
            properties.load(fis);
            authToken = properties.getProperty("auth_token");
            categoriesUrl = properties.getProperty("categories_url");
            System.out.println("Properties loaded successfully.");
        } catch (IOException e) {
            System.err.println("Error loading properties file: " + e.getMessage());
        }
    }

    private static void processExcelData() {
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            boolean isHeader = true;
            for (Row row : sheet) {
                if (isHeader) {
                    isHeader = false; // Skip header row
                    continue;
                }
                Cell curlCell = row.getCell(2); // Assuming cURL command is in column C (index 2)
                Cell responseCell = row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); // Response in column F (index 5)
                
                if (curlCell != null) {
                    String curlCommand = curlCell.getStringCellValue();
                    System.out.println("Processing row " + row.getRowNum() + " - cURL: " + curlCommand);
                    
                    if (!isValidCurlFormat(curlCommand)) {
                        System.err.println("Invalid cURL format detected at row " + row.getRowNum());
                        responseCell.setCellValue("Invalid cURL format");
                    } else {
                        String formattedCurl = formatCurl(curlCommand);
                        System.out.println("Formatted cURL: " + formattedCurl);
                        String response = executeApiRequest(formattedCurl);
                        responseCell.setCellValue(response);
                        System.out.println("API response written for row " + row.getRowNum());
                    }
                }
            }
            try (FileOutputStream fos = new FileOutputStream(excelFilePath)) {
                workbook.write(fos);
                System.out.println("Excel file updated successfully.");
            }
        } catch (IOException e) {
            System.err.println("Error processing Excel file: " + e.getMessage());
        }
    }

    private static boolean isValidCurlFormat(String curlCommand) {
        // Check if the command starts with 'curl' and contains a URL
        return curlCommand.startsWith("curl ") && curlCommand.contains("http");
    }

    private static String formatCurl(String curlCommand) {
        // Replace placeholders with actual values from the properties file
        if (curlCommand.contains("{{categories_url}}")) {
            curlCommand = curlCommand.replace("{{categories_url}}", categoriesUrl);
        }
        if (curlCommand.contains("{{auth_token}}")) {
            curlCommand = curlCommand.replace("{{auth_token}}", authToken);
        }
        return curlCommand;
    }

    private static String executeApiRequest(String curlCommand) {
        try {
            RequestSpecification request = RestAssured.given();

            // Extract URL
            Pattern urlPattern = Pattern.compile("'(https?://[^']+)'");
            Matcher urlMatcher = urlPattern.matcher(curlCommand);
            String url = urlMatcher.find() ? urlMatcher.group(1) : "";

            // Extract headers
            Pattern headerPattern = Pattern.compile("--header '([^:]+): ([^']+)'");
            Matcher headerMatcher = headerPattern.matcher(curlCommand);
            while (headerMatcher.find()) {
                request.header(headerMatcher.group(1), headerMatcher.group(2));
            }

            // Extract body data
            Pattern dataPattern = Pattern.compile("--data '([^']+)'");
            Matcher dataMatcher = dataPattern.matcher(curlCommand);
            String requestBody = dataMatcher.find() ? dataMatcher.group(1) : "";

            // Determine HTTP method dynamically
            String method = "GET"; // Default method is GET
            if (curlCommand.contains("--data")) {
                method = "POST"; // If data is present, it's likely a POST request
            }

            // Print the final request details
            System.out.println("Request URL: " + url);
            System.out.println("Request Headers: ");
            headerMatcher.reset();
            while (headerMatcher.find()) {
                System.out.println(headerMatcher.group(1) + ": " + headerMatcher.group(2));
            }
            System.out.println("Request Body: " + requestBody);
            System.out.println("HTTP Method: " + method);

            Response response;

            // Handle different HTTP methods
            switch (method) {
                case "POST":
                    response = request.body(requestBody).post(url);
                    System.out.println("POST request executed");
                    break;
                case "PUT":
                    response = request.body(requestBody).put(url);
                    System.out.println("PUT request executed");
                    break;
                case "DELETE":
                    response = request.delete(url);
                    System.out.println("DELETE request executed");
                    break;
                case "PATCH":
                    response = request.body(requestBody).patch(url);
                    System.out.println("PATCH request executed");
                    break;
                default:
                    // Default to GET request
                    response = request.get(url);
                    System.out.println("GET request executed");
                    break;
            }

            // Print the response
            System.out.println("API Response Status Code: " + response.getStatusCode());
            System.out.println("API Response Body: " + response.getBody().asString());
            return response.getBody().asString();
        } catch (Exception e) {
            System.err.println("Error executing API request: " + e.getMessage());
            return "Error executing API request: " + e.getMessage();
        }
    }
}










