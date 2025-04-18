package postmanCollectionsHelper;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

public class JsonBuilderSample {

    public static void main(String[] args) throws Exception {
        ObjectMapper mapper = new ObjectMapper();

        // Root JSON object
        ObjectNode root = mapper.createObjectNode();

        // Add properties
        root.put("name", "John Doe");
        root.put("age", 30);

        // Create a nested JSON object
        ObjectNode addressNode = mapper.createObjectNode();
        addressNode.put("street", "123 Main St");
        addressNode.put("city", "Anytown");
        addressNode.put("zip", "12345");

        // Attach nested object
        root.set("address", addressNode);

        // Create an array node
        ArrayNode phoneNumbersNode = mapper.createArrayNode();

        ObjectNode phoneNumber1 = mapper.createObjectNode();
        phoneNumber1.put("type", "home");
        phoneNumber1.put("number", "123-456-7890");

        ObjectNode phoneNumber2 = mapper.createObjectNode();
        phoneNumber2.put("type", "work");
        phoneNumber2.put("number", "987-654-3210");

        // Add objects to array
        phoneNumbersNode.add(phoneNumber1);
        phoneNumbersNode.add(phoneNumber2);

        // Attach array to root
        root.set("phoneNumbers", phoneNumbersNode);

        // Convert to JSON string and print
        String jsonString = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(root);
        System.out.println(jsonString);
    }
}




