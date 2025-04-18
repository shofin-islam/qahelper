package postmanCollectionsHelper;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class PostmanCollectionMerger {

    private static int totalProcessed = 0;
    private static int totalSkipped = 0;
    private static int totalAdded = 0;

    public static void main(String[] args) {
        String inputDir = "/Users/bs00880/myworkspace/collections";
        String outputDir = "/Users/bs00880/myworkspace/output";
        System.out.println("******** PostmanCollectionMerger -> Execution Started");
        File dir = new File(inputDir);
        File outputDirectory = new File(outputDir);
        if (!outputDirectory.exists()) {
            outputDirectory.mkdirs();
        }

        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        String dynamicFolderName = "Merged_Collection_" + timestamp;
        File finalOutputDir = new File(outputDirectory, dynamicFolderName);
        finalOutputDir.mkdirs();

        ObjectMapper objectMapper = new ObjectMapper();
        File[] files = dir.listFiles((d, name) -> name.endsWith(".json"));

        Set<String> processedRequests = new HashSet<>();
        ArrayNode mergedFolders = objectMapper.createArrayNode();

        if (files != null) {
            System.out.println("Processing files from directory: " + inputDir);
            for (File file : files) {
                try {
                    System.out.println("📂 Processing file: " + file.getName());
                    JsonNode collection = objectMapper.readTree(file);
                    ArrayNode items = (ArrayNode) collection.get("item");

                    if (items != null) {
                        for (JsonNode item : items) {
                            processItem(item, processedRequests, mergedFolders, "");
                        }
                    }
                } catch (IOException e) {
                    System.err.println("❌ Error reading file: " + file.getName());
                    e.printStackTrace();
                }
            }
        }

        System.out.println("\n===== ✅ MERGE SUMMARY ✅ =====");
        System.out.println("📌 Total Requests Processed: " + totalProcessed);
        System.out.println("🚫 Total Requests Skipped (Duplicates): " + totalSkipped);
        System.out.println("✅ Total Requests Added to Merged Collection: " + totalAdded);

        ObjectNode finalCollection = objectMapper.createObjectNode();
        finalCollection.putObject("info")
            .put("_postman_id", UUID.randomUUID().toString())
            .put("name", "Merged API Collection")
            .put("schema", "https://schema.getpostman.com/json/collection/v2.1.0/collection.json");
        finalCollection.set("item", mergedFolders);

        saveCollection(finalOutputDir, "merged_collection.json", finalCollection, objectMapper);
    }

    private static void processItem(JsonNode item, Set<String> processedRequests, ArrayNode mergedFolders, String folderPath) {
        if (item.has("item")) {
            ObjectNode folder = (ObjectNode) item.deepCopy();
            ArrayNode subItems = (ArrayNode) folder.get("item");
            ArrayNode newSubItems = new ObjectMapper().createArrayNode();

            for (JsonNode subItem : subItems) {
                processItem(subItem, processedRequests, newSubItems, folderPath + "/" + item.get("name").asText());
            }

            if (!newSubItems.isEmpty()) {
                folder.set("item", newSubItems);
                mergedFolders.add(folder);
                System.out.println("📂 Added folder: " + folderPath + "/" + item.get("name").asText());
            }
        } else {
            JsonNode request = item.get("request");
            if (request != null) {
                totalProcessed++; // Every request gets counted
                String method = request.get("method").asText();
                JsonNode urlNode = request.get("url");
                if (urlNode != null) {
                    String url = urlNode.get("raw").asText();
                    String uniqueKey = method + "_" + url;

                    if (!processedRequests.contains(uniqueKey)) {
                        processedRequests.add(uniqueKey);
                        mergedFolders.add(item);
                        totalAdded++;
                        System.out.println("✅ Added: " + method + " " + url);
                    } else {
                        totalSkipped++;
                        System.out.println("🚫 Skipped Duplicate: " + method + " " + url);
                    }
                }
            }
        }
    }

    private static void saveCollection(File outputDir, String fileName, ObjectNode collection, ObjectMapper objectMapper) {
        try {
            File outputFile = new File(outputDir, fileName);
            objectMapper.writerWithDefaultPrettyPrinter().writeValue(outputFile, collection);
            System.out.println("💾 Saved collection to: " + outputFile.getAbsolutePath());
        } catch (IOException e) {
            System.err.println("❌ Error saving file: " + fileName);
            e.printStackTrace();
        }
    }
}
