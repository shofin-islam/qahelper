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
        String inputFilePath = "MS-Wise-Api-Validation-Test-Results-Automate.xlsx";
        String outputFilePath = getTimestampedFileName("MS-Wise-Api-Validation-Test-Results-Automate.xlsx");

        compareJsonWithKeyMatching(inputFilePath, outputFilePath);
    }

    

    public static void compareJsonWithKeyMatching(String inputFilePath, String outputFilePath) throws Exception {
        FileInputStream fis = new FileInputStream(inputFilePath);
        Workbook inputWorkbook = new XSSFWorkbook(fis);
        Sheet inputSheet = inputWorkbook.getSheetAt(4);

        Map<String, List<String>> collectionA = new LinkedHashMap<>();
        Map<String, List<String>> collectionB = new LinkedHashMap<>();

        // Load data into collections
        for (int i = 1; i <= inputSheet.getLastRowNum(); i++) {
            Row row = inputSheet.getRow(i);
            if (row == null) continue;

            // A_Key = C|D|Q -> (2, 3, 16)
            String colC = cleanCellValue(row.getCell(2));
            String colD = cleanCellValue(row.getCell(3));
            String colQ = cleanCellValue(row.getCell(16));
            String aKey = colC + "|" + colD + "|" + colQ;

            List<String> aValues = new ArrayList<>();
            for (int ii = 0; ii <= 17; ii++) {
                aValues.add(getCellValueAsString(row.getCell(ii)));
            }
            collectionA.put(aKey, aValues);
            System.out.println("A_Key: " + aKey);

            // B_Key = S|T|Z -> (18, 19, 25)
            String colS = cleanCellValue(row.getCell(18));
            String colT = cleanCellValue(row.getCell(19));
            String colZ = cleanCellValue(row.getCell(25));
            String bKey = colS + "|" + colT + "|" + colZ;

            List<String> bValues = new ArrayList<>();
            for (int ij = 18; ij <= 25; ij++) {
                bValues.add(getCellValueAsString(row.getCell(ij)));
            }
            collectionB.put(bKey, bValues);
            System.out.println("B_Key: " + bKey);
        }

        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet masterSheet = outputWorkbook.createSheet("Master");
        Sheet collectionASheet = outputWorkbook.createSheet("CollectionA");
        Sheet collectionBSheet = outputWorkbook.createSheet("CollectionB");

        // Write Collection A
        int aRowNum = 0;
        for (Map.Entry<String, List<String>> entry : collectionA.entrySet()) {
            Row row = collectionASheet.createRow(aRowNum++);
            int cellNum = 0;
            for (String val : entry.getValue()) {
                row.createCell(cellNum++).setCellValue(val);
            }
        }

        // Write Collection B
        int bRowNum = 0;
        for (Map.Entry<String, List<String>> entry : collectionB.entrySet()) {
            Row row = collectionBSheet.createRow(bRowNum++);
            int cellNum = 0;
            for (String val : entry.getValue()) {
                row.createCell(cellNum++).setCellValue(val);
            }
        }

        // Master comparison and output
        int mRowNum = 0;
        for (Map.Entry<String, List<String>> aEntry : collectionA.entrySet()) {
            Row mRow = masterSheet.createRow(mRowNum++);
            int cellNum = 0;
            for (String val : aEntry.getValue()) {
                mRow.createCell(cellNum++).setCellValue(val);
            }

            String aKey = aEntry.getKey();
            if (collectionB.containsKey(aKey)) {
                List<String> bVals = collectionB.get(aKey);
                for (String val : bVals) {
                    mRow.createCell(cellNum++).setCellValue(val);
                }

                // JSON comparison: H (index 7 in A) and V (index 21 in full sheet, so 21 - 18 = 3 in B)
                String jsonA = aEntry.getValue().size() > 7 ? aEntry.getValue().get(7) : "";
                String jsonB = bVals.size() > 3 ? bVals.get(3) : "";

                System.out.println("Comparing A_Key: " + aKey);
//                System.out.println("  JSON A: " + jsonA);
//                System.out.println("  JSON B: " + jsonB);

                String comparison = compareJson(jsonA, jsonB);
                mRow.createCell(26).setCellValue(comparison);
            } else {
                mRow.createCell(26).setCellValue("Not Found in Collection B");
                System.out.println("Not Found in B: " + aKey);
            }
        }

        FileOutputStream fos = new FileOutputStream(outputFilePath);
        outputWorkbook.write(fos);
        fos.close();
        inputWorkbook.close();
        outputWorkbook.close();

        System.out.println("Comparison complete. Output written to: " + outputFilePath);
    }


    private static String compareJson(String jsonStr1, String jsonStr2) {
        try {
            JsonNode jsonNode1 = objectMapper.readTree(jsonStr1);
            JsonNode jsonNode2 = objectMapper.readTree(jsonStr2);
            List<String> differences = new ArrayList<>();
            compareJsonNodes(jsonNode1, jsonNode2, "", differences);
            return differences.isEmpty() ? "Identical" : String.join("; ", differences);
        } catch (Exception e) {
            return "Invalid JSON: " + e.getMessage();
        }
    }

    private static void compareJsonNodes(JsonNode node1, JsonNode node2, String path, List<String> differences) {
        if (!node1.getNodeType().equals(node2.getNodeType())) {
            differences.add("Type mismatch at " + path);
            return;
        }
        if (node1.isObject()) {
            Iterator<Map.Entry<String, JsonNode>> fields = node1.fields();
            while (fields.hasNext()) {
                Map.Entry<String, JsonNode> entry = fields.next();
                String key = entry.getKey();
                if (!node2.has(key)) {
                    differences.add("Missing key in JSON2 at " + path + "." + key);
                    continue;
                }
                compareJsonNodes(entry.getValue(), node2.get(key), path + "." + key, differences);
            }
            Iterator<Map.Entry<String, JsonNode>> fields2 = node2.fields();
            while (fields2.hasNext()) {
                String key = fields2.next().getKey();
                if (!node1.has(key)) {
                    differences.add("Missing key in JSON1 at " + path + "." + key);
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

    private static String cleanCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                double val = cell.getNumericCellValue();
                if (val == (long) val)
                    return String.valueOf((long) val);
                else
                    return String.valueOf(val);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula().trim();
            default:
                return "";
        }
    }
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA: return cell.getCellFormula();
            default: return "";
        }
    }
}
