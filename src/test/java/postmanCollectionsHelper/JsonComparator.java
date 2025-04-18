package postmanCollectionsHelper;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class JsonComparator {

    private static final ObjectMapper objectMapper = new ObjectMapper();

    public static void main(String[] args) throws Exception {
        String inputFilePath = "API Responses Validation - PHP Version Upgrage M1.xlsx";   // Input Excel file
        String outputFilePath = getTimestampedFileName("API Responses Validation - PHP Version Upgrage M1.xlsx");

        compareJsonFromExcel(inputFilePath, outputFilePath);
    }

    public static void compareJsonFromExcel(String inputFilePath, String outputFilePath) throws Exception {
        FileInputStream fis = new FileInputStream(inputFilePath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0); // Read first sheet

        System.out.println(sheet.getSheetName());

        // Ensure header exists
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) {
            headerRow = sheet.createRow(0);
        }

        // Set header for comparison result at index 10
        Cell headerCell = headerRow.getCell(10);
        if (headerCell == null) {
            headerCell = headerRow.createCell(10, CellType.STRING);
            headerCell.setCellValue("Comparison Result");
        }

        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            Cell json1Cell = row.getCell(6); // JSON 1 from Column 5
            Cell json2Cell = row.getCell(7); // JSON 2 from Column 7
            
            

            if (json1Cell == null || json2Cell == null) {
                continue;
            }

            String json1 = getCellValueAsString(json1Cell);
            String json2 = getCellValueAsString(json2Cell);

            String comparisonResult = compareJson(json1, json2);

            // Write comparison result in column 10 (Index 10) of the same row
            Cell resultCell = row.createCell(10, CellType.STRING);
            resultCell.setCellValue(comparisonResult);
        }

        fis.close();

        // Save output Excel
        FileOutputStream fos = new FileOutputStream(outputFilePath);
        workbook.write(fos);
        fos.close();
        workbook.close();

        System.out.println("Comparison summary saved in: " + outputFilePath);
    }

    private static String compareJson(String jsonStr1, String jsonStr2) {
        try {
            JsonNode jsonNode1 = objectMapper.readTree(jsonStr1);
            JsonNode jsonNode2 = objectMapper.readTree(jsonStr2);

            List<String> differences = new ArrayList<>();
            compareJsonNodes(jsonNode1, jsonNode2, "", differences);

            return differences.isEmpty() ? "Identical" : String.join("; ", differences);
        } catch (Exception e) {
            return "Invalid JSON format: " + e.getMessage();
        }
    }

    private static void compareJsonNodes(JsonNode node1, JsonNode node2, String path, List<String> differences) {
        if (!node1.getNodeType().equals(node2.getNodeType())) {
            differences.add("Type mismatch at " + path + " - " + node1.getNodeType() + " vs " + node2.getNodeType());
            return;
        }

        if (node1.isObject()) {
            Iterator<Map.Entry<String, JsonNode>> fields = node1.fields();
            while (fields.hasNext()) {
                Map.Entry<String, JsonNode> entry = fields.next();
                String key = entry.getKey();
                if (!node2.has(key)) {
                    differences.add("Missing key in JSON 2 at " + path + "." + key);
                    continue;
                }
                compareJsonNodes(entry.getValue(), node2.get(key), path + "." + key, differences);
            }

            Iterator<Map.Entry<String, JsonNode>> fields2 = node2.fields();
            while (fields2.hasNext()) {
                Map.Entry<String, JsonNode> entry = fields2.next();
                String key = entry.getKey();
                if (!node1.has(key)) {
                    differences.add("Missing key in JSON 1 at " + path + "." + key);
                }
            }
        } else if (node1.isArray()) {
            if (node1.size() != node2.size()) {
                differences.add("Array size mismatch at " + path);
                return;
            }
            for (int i = 0; i < node1.size(); i++) {
                compareJsonNodes(node1.get(i), node2.get(i), path + "[" + i + "]", differences);
            }
        } else {
            if (!node1.asText().equals(node2.asText())) {
                differences.add("Value mismatch at " + path + " - " + node1.asText() + " vs " + node2.asText());
            }
        }
    }

    private static String getTimestampedFileName(String baseFileName) {
        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        return baseFileName.replace(".xlsx", "_" + timestamp + ".xlsx");
    }
    
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula(); // Handle formula cells if needed
            case BLANK:
                return "";
            default:
                return "";
        }
    }

}
