package postmanCollectionsHelper;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

public class UniqueSortedPostmanMerger {

    public static void main(String[] args) {
        String inputDir = "/Users/bs00880/myworkspace/collections";
        String outputDir = "/Users/bs00880/myworkspace/output";
        
        File dir = new File(inputDir);
        File outputDirectory = new File(outputDir);
        if (!outputDirectory.exists()) {
            outputDirectory.mkdirs();
        }

        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        String outputFileName = "Unique_Sorted_Collection_" + timestamp + ".json";
        File outputFile = new File(outputDirectory, outputFileName);

        ObjectMapper objectMapper = new ObjectMapper();
        File[] files = dir.listFiles((d, name) -> name.endsWith(".json"));

        if (files == null) {
            System.out.println("No Postman collections found in directory: " + inputDir);
            return;
        }

        int totalProcessed = 0;
        Set<String> uniqueRequests = new HashSet<>();
        List<JsonNode> sortedRequests = new ArrayList<>();

        System.out.println("Processing collections from directory: " + inputDir);
        for (File file : files) {
            try {
                System.out.println("Processing file: " + file.getName());
                JsonNode collection = objectMapper.readTree(file);
                ArrayNode items = (ArrayNode) collection.get("item");

                if (items != null) {
                    totalProcessed += extractRequests(items, uniqueRequests, sortedRequests);
                }
            } catch (IOException e) {
                System.err.println("Error reading file: " + file.getName());
                e.printStackTrace();
            }
        }

        // Sort requests alphabetically by URL
        sortedRequests.sort(Comparator.comparing(req -> req.get("request").get("url").get("raw").asText()));

        // Create final collection JSON
        ObjectNode finalCollection = objectMapper.createObjectNode();
        finalCollection.putObject("info")
            .put("_postman_id", UUID.randomUUID().toString())
            .put("name", "Unique Sorted API Collection")
            .put("schema", "https://schema.getpostman.com/json/collection/v2.1.0/collection.json");
        finalCollection.set("item", objectMapper.valueToTree(sortedRequests));

        // Save the collection
        try {
            objectMapper.writerWithDefaultPrettyPrinter().writeValue(outputFile, finalCollection);
            System.out.println("✅ Unique Sorted Collection saved at: " + outputFile.getAbsolutePath());
        } catch (IOException e) {
            System.err.println("Error saving file: " + outputFile.getAbsolutePath());
            e.printStackTrace();
        }

        // Print final summary
        int totalUnique = uniqueRequests.size();
        int totalSkipped = totalProcessed - totalUnique;
        System.out.println("\n===== UNIQUE SORTED COLLECTION SUMMARY =====");
        System.out.println("Total Requests Processed: " + totalProcessed);
        System.out.println("Total Unique Requests: " + totalUnique);
        System.out.println("Total Skipped Requests (Duplicates): " + totalSkipped);
    }

    private static int extractRequests(ArrayNode items, Set<String> uniqueRequests, List<JsonNode> sortedRequests) {
        int count = 0;
        for (JsonNode item : items) {
            if (item.has("item")) {
                // Recursively process folders
                count += extractRequests((ArrayNode) item.get("item"), uniqueRequests, sortedRequests);
            } else {
                // Process an actual request
                JsonNode request = item.get("request");
                if (request != null) {
                    count++;
                    String method = request.get("method").asText();
                    JsonNode urlNode = request.get("url");

                    if (urlNode != null) {
                        String url = urlNode.get("raw").asText();
                        String uniqueKey = method + "_" + url;

                        if (!uniqueRequests.contains(uniqueKey)) {
                            uniqueRequests.add(uniqueKey);
                            sortedRequests.add(item);
                            System.out.println("Added unique request: " + method + " " + url);
                        } else {
                            System.out.println("Skipped duplicate request: " + method + " " + url);
                        }
                    }
                }
            }
        }
        return count;
    }
}
