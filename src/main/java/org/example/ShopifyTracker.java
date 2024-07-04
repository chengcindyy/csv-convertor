package org.example;

import io.github.cdimascio.dotenv.Dotenv;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ShopifyTracker {

    private final String shopifyUrl;
    private final String accessToken;
    private static final Logger LOGGER = Logger.getLogger(ShopifyTracker.class.getName());
    private static final String TRACKER_URI = "http://www.t-cat.com.tw/Inquire/Trace.aspx?no=";
    private static final String TRACKER_CARRIER = "Other";

    public ShopifyTracker() {
        Dotenv dotenv = Dotenv.load();
        this.shopifyUrl = dotenv.get("SHOPIFY_URL");
        this.accessToken = dotenv.get("SHOPIFY_ACCESS_TOKEN");
    }

    public static void ImportDataProcessor(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assume data is in the first sheet
            for (Row row : sheet) {
                // Extract data from each row
                String orderId = row.getCell(0).getStringCellValue();
                String trackingNumber = row.getCell(2).getStringCellValue();

                ShopifyTracker tracker = new ShopifyTracker();
                int responseCode = tracker.updateTrackingInfo(orderId, trackingNumber, TRACKER_URI, TRACKER_CARRIER);
                if (responseCode == HttpURLConnection.HTTP_CREATED) {
                    System.out.println("Data uploaded to Shopify successfully for order " + orderId);
                } else {
                    System.err.println("Failed to upload data to Shopify for order " + orderId + ". Response code: " + responseCode);
                }
            }
        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, "Xlsx file process error", e);
        }
    }

    public int updateTrackingInfo(String orderId, String trackingNumber, String trackingUrl, String carrier) throws IOException {
        String endpoint = String.format("%s/admin/api/2021-07/orders/%s/fulfillments.json", shopifyUrl, orderId);
        System.out.println(endpoint);
        URL url = new URL(endpoint);
        HttpURLConnection connection = createConnection(url);

        String jsonPayload;
        if (carrier.equalsIgnoreCase("Other")) {
            jsonPayload = String.format(
                    "{\"fulfillment\": {\"tracking_number\": \"%s\", \"tracking_url\": \"%s\", \"tracking_company\": \"%s\"}}",
                    trackingNumber, trackingUrl, carrier);
        } else {
            jsonPayload = String.format(
                    "{\"fulfillment\": {\"tracking_number\": \"%s\", \"tracking_url\": \"%s\", \"tracking_company\": \"%s\"}}",
                    trackingNumber, trackingUrl, carrier);
        }

        try (OutputStream os = connection.getOutputStream()) {
            byte[] input = jsonPayload.getBytes(StandardCharsets.UTF_8);
            os.write(input, 0, input.length);
        }

        return connection.getResponseCode();
    }


    private HttpURLConnection createConnection(URL url) throws IOException {
        System.out.println("Connecting to " + url.toString());
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json");
        connection.setRequestProperty("X-Shopify-Access-Token", accessToken);
        connection.setDoOutput(true);
        System.out.println("connected to: "+connection.getHeaderFields());
        return connection;
    }

    public static void main(String[] args) {
        ShopifyTracker tracker = new ShopifyTracker();
        File file = new File("path/to/your/excel/file.xlsx");
        tracker.ImportDataProcessor(file);
    }
}
