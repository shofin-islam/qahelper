package postmanCollectionsHelper;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.*;
import java.util.*;

public class ExcelMergerWithSheetTraceAndDuplicates {

    private static final int MAX_CELL_LENGTH = 32767;
    private static final String OUTPUT_FILE = "MergedTestCases.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        String folderPath = "/Users/bs00880/myworkspace/excels"; // Change this path
        File folder = new File(folderPath);
        File[] files = folder.listFiles((dir, name) -> name.endsWith(".xlsx"));

        if (files == null || files.length == 0) {
            System.out.println("No Excel files found in the folder.");
            return;
        }

        // Use streaming SXSSFWorkbook to reduce memory usage
        SXSSFWorkbook mergedWorkbook = new SXSSFWorkbook(100); // Keep 100 rows in memory
        Sheet mergedSheet = mergedWorkbook.createSheet("MergedData");
        int mergedRowNum = 0;

        CellStyle headerStyle = mergedWorkbook.createCellStyle();
        Font headerFont = mergedWorkbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);

        CellStyle redHighlightStyle = mergedWorkbook.createCellStyle();
        redHighlightStyle.cloneStyleFrom(headerStyle);
        redHighlightStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        redHighlightStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Set<String> allSheetNames = new HashSet<>();
        Map<String, List<String>> duplicateSheetSources = new HashMap<>();

        for (File file : files) {
            try (Workbook workbook = WorkbookFactory.create(file)) {
                String fileName = file.getName();

                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    String sheetName = sheet.getSheetName();
                    boolean isDuplicateSheet = !allSheetNames.add(sheetName);

                    if (isDuplicateSheet) {
                        duplicateSheetSources.computeIfAbsent(sheetName, k -> new ArrayList<>()).add(fileName);
                    }

                    Iterator<Row> rowIterator = sheet.iterator();
                    boolean isFirstRow = true;

                    while (rowIterator.hasNext()) {
                        Row row = rowIterator.next();

                        // Skip rows if all cells (excluding first two) are empty
                        boolean isDataRowEmpty = true;
                        for (int col = 0; col < row.getLastCellNum(); col++) {
                            Cell cell = row.getCell(col);
                            if (cell != null && cell.getCellType() != CellType.BLANK) {
                                if (cell.getCellType() == CellType.STRING && !cell.getStringCellValue().trim().isEmpty()) {
                                    isDataRowEmpty = false;
                                    break;
                                } else if (cell.getCellType() != CellType.STRING) {
                                    isDataRowEmpty = false;
                                    break;
                                }
                            }
                        }

                        if (isDataRowEmpty) continue; // skip empty row

                        Row newRow = mergedSheet.createRow(mergedRowNum++);

                        // Sheet name cell
                        Cell sheetNameCell = newRow.createCell(0);
                        sheetNameCell.setCellValue(sheetName);
                        if (isDuplicateSheet) sheetNameCell.setCellStyle(redHighlightStyle);

                        // File name cell
                        Cell fileNameCell = newRow.createCell(1);
                        fileNameCell.setCellValue(fileName);

                        for (int col = 0; col < row.getLastCellNum(); col++) {
                            Cell oldCell = row.getCell(col);
                            Cell newCell = newRow.createCell(col + 2);

                            if (oldCell != null) {
                                switch (oldCell.getCellType()) {
                                    case STRING:
                                        String val = oldCell.getStringCellValue();
                                        if (val.length() > MAX_CELL_LENGTH)
                                            val = val.substring(0, MAX_CELL_LENGTH - 3) + "...";
                                        newCell.setCellValue(val);
                                        break;
                                    case NUMERIC:
                                        newCell.setCellValue(oldCell.getNumericCellValue());
                                        break;
                                    case BOOLEAN:
                                        newCell.setCellValue(oldCell.getBooleanCellValue());
                                        break;
                                    default:
                                        newCell.setCellValue("");
                                }
                            }
                        }

                        if (isFirstRow) {
                            newRow.setRowStyle(headerStyle);
                            isFirstRow = false;
                        }
                    }

                    System.out.println("✅ Processed: " + sheetName + " from " + fileName);
                }
            }
        }

        try (FileOutputStream outStream = new FileOutputStream(OUTPUT_FILE)) {
            mergedWorkbook.write(outStream);
            System.out.println("✅ Merged data written to " + OUTPUT_FILE);
        }

        mergedWorkbook.dispose(); // clean up temporary files
        mergedWorkbook.close();

        // Write duplicate sheet info
        if (!duplicateSheetSources.isEmpty()) {
            SXSSFWorkbook dupWorkbook = new SXSSFWorkbook();
            Sheet dupSheet = dupWorkbook.createSheet("DuplicateSheets");

            Row dupHeader = dupSheet.createRow(0);
            dupHeader.createCell(0).setCellValue("Sheet Name");
            dupHeader.createCell(1).setCellValue("Source File Name");

            int rowIdx = 1;
            for (Map.Entry<String, List<String>> entry : duplicateSheetSources.entrySet()) {
                for (String fileName : entry.getValue()) {
                    Row row = dupSheet.createRow(rowIdx++);
                    row.createCell(0).setCellValue(entry.getKey());
                    row.createCell(1).setCellValue(fileName);
                }
            }

            try (FileOutputStream dupOut = new FileOutputStream("DuplicateSheetNames.xlsx")) {
                dupWorkbook.write(dupOut);
                System.out.println("📄 Duplicate sheet names exported to DuplicateSheetNames.xlsx");
            }

            dupWorkbook.dispose();
            dupWorkbook.close();
        }
    }
}
