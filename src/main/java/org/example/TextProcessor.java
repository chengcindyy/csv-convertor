package org.example;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.*;

public class TextProcessor {

    private static final Logger LOGGER = Logger.getLogger(Main.class.getName());
    private static final List<String> CUSTOM_HEADERS = List.of("Order No.", "Email", "SKU", "Discount Code", "Product", "Qty", "Total Price", "Customer", "Address", "Country", "Phone", "Notes", "Fulfillment Note");
    private static final List<String> NEW_HEADERS = List.of("出貨日", "指定配達日", "訂單編號", "代收金額", "品項", "備註", "收件人", "手機(收)", "電話(收)", "地址(收)", "契客代號", "溫層", "規格", "配送時間帶", "會員編號", "寄件人", "寄件人地址", "寄件人電話", "複數件", "數量檢查/可刪除", "Email/可刪除", "訂單數/可刪除", "合併狀態/可刪除");

    public void CSVToExcelProcessor(File file, String outputFilePath) {
        try (FileReader reader = new FileReader(file); CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT.builder().setHeader().setSkipHeaderRecord(true).build())) {

            Map<String, Orders> ordersMap = processCSVRecords(csvParser);

            // Create two workbooks for TW and LOCAL
            Workbook workbookTW = new XSSFWorkbook();
            Workbook workbookLocal = new XSSFWorkbook();

            initializeWorkbooks(workbookTW);
            initializeWorkbooks(workbookLocal);

            Map<String, Integer> rowNumsTW = new HashMap<>();
            Map<String, Integer> rowNumsLocal = new HashMap<>();
            rowNumsTW.put("Cancelled Orders", 1);
            rowNumsLocal.put("Cancelled Orders", 1);
            rowNumsTW.put("Unfulfilled Orders", 1);
            rowNumsLocal.put("Unfulfilled Orders", 1);

            writeDataToSheets(ordersMap, workbookTW, rowNumsTW, "TW");
            writeDataToSheets(ordersMap, workbookLocal, rowNumsLocal, "LOCAL");

            // Create new sheets with NEW_HEADERS
            writeDataToNewSheets(ordersMap, workbookTW, "TW");
            writeDataToNewSheets(ordersMap, workbookLocal, "LOCAL");

            saveWorkbook(workbookTW, outputFilePath + File.separator + file.getName().replace(".csv", "_TW.xlsx"));
            saveWorkbook(workbookLocal, outputFilePath + File.separator + file.getName().replace(".csv", "_LOCAL.xlsx"));

            System.out.println("Processed files saved as: " + outputFilePath);

        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, "Csv file process error", e);
        }
    }

    private Map<String, Orders> processCSVRecords(CSVParser csvParser) throws IOException {
        Map<String, Orders> ordersMap = new HashMap<>();
        Map<String, String> emailToAddressMap = new HashMap<>();
        Map<String, List<String>> fulfillmentNotesMap = new HashMap<>();

        for (CSVRecord record : csvParser) {
            String orderNumber = record.get("Name");
            String email = record.get("Email").trim();
            String product = record.get("Lineitem name").trim();

            String city = record.get("Shipping City").trim();
            String fullAddress = record.get("Shipping Street").trim();
            if (fullAddress.contains(city)) {
                fullAddress = fullAddress.replace(city, "").trim();
            }
            String address = city + fullAddress;

            String note = record.get("Notes").trim();
            int currentQuantity = Integer.parseInt(record.get("Lineitem quantity").trim());
            String totalPriceStr = record.get("Total").trim();
            String discountAmountStr = record.get("Discount Amount").trim();
            String cancelledAt = record.get("Cancelled at").trim();
            String fulfillmentStatus = record.get("Fulfillment Status").trim();
            String fulfilledAt = record.get("Fulfilled at");

            double totalPrice = totalPriceStr.isEmpty() ? 0.0 : Double.parseDouble(totalPriceStr);
            double discountAmount = discountAmountStr.isEmpty() ? 0.0 : Double.parseDouble(discountAmountStr);

            if (!address.isEmpty()) {
                emailToAddressMap.put(email, address);
            }

            // Create a key for fulfillment note processing
            String fulfillmentNoteKey = email + "|" + address;
            fulfillmentNotesMap.putIfAbsent(fulfillmentNoteKey, new ArrayList<>());
            fulfillmentNotesMap.get(fulfillmentNoteKey).add(orderNumber + ": " + currentQuantity);

            // Create a unique key for orders
            String orderKey = orderNumber + "|" + email + "|" + product + "|" + address;
            Orders order = ordersMap.getOrDefault(orderKey, new Orders());

            order.setOrderNo(orderNumber);
            order.setProduct(processProductName(product));
            order.setQuantity(order.getQuantity() + currentQuantity);
            order.setEmail(email);
            order.setAddress(address);
            order.setCountry(record.get("Shipping Country"));
            order.setDisCountCode(record.get("Discount Code"));

            if (totalPrice == 0.0) {
                order.setTotalPrice(discountAmount);
            } else {
                order.setTotalPrice(totalPrice);
            }

            order.setName(record.get("Shipping Name"));
            order.setPhone(record.get("Shipping Phone"));
            order.setSku(record.get("Lineitem sku"));
            order.setCancelledAt(cancelledAt);
            order.setFulfillmentStatus(fulfillmentStatus);
            order.setFulfillmentAt(fulfilledAt);

            if (!note.isEmpty()) {
                order.getNotes().add(note);
            }

            ordersMap.put(orderKey, order);
        }

        // Now assign fulfillment notes to the orders
        for (Orders order : ordersMap.values()) {
            String fulfillmentNoteKey = order.getEmail() + "|" + order.getAddress();
            if (fulfillmentNotesMap.containsKey(fulfillmentNoteKey)) {
                order.setFulfillmentNote(String.join("; ", fulfillmentNotesMap.get(fulfillmentNoteKey)));
                if (fulfillmentNotesMap.get(fulfillmentNoteKey).size() > 1) {
                    order.setMerged(true);
                }
            } else {
                order.setFulfillmentNote("");  // Handle cases where there is no fulfillment note
            }
        }

        for (Orders order : ordersMap.values()) {
            if (order.getAddress().isEmpty() && emailToAddressMap.containsKey(order.getEmail())) {
                order.setAddress(emailToAddressMap.get(order.getEmail()));
            }
        }

        return ordersMap;
    }

    private String processProductName(String product) {
        Pattern pattern = Pattern.compile("(\\d+\\.?\\d?R).*(\\d+KG\\s*/\\s*箱).*(第.*團)");
        Matcher matcher = pattern.matcher(product);

        if (matcher.find()) {
            String productCode = matcher.group(1);
            String weight = matcher.group(2);
            String group = matcher.group(3);

            return productCode + " " + weight + " - " + group;
        }
        return product;
    }

    private void initializeWorkbooks(Workbook workbook) {
        Sheet sheet = workbook.createSheet("Cancelled Orders");
        Row headerRow = sheet.createRow(0);
        for (int i = 0; TextProcessor.CUSTOM_HEADERS != null && i < TextProcessor.CUSTOM_HEADERS.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(TextProcessor.CUSTOM_HEADERS.get(i));
        }
    }

    private void writeDataToSheets(Map<String, Orders> ordersMap, Workbook workbook, Map<String, Integer> rowNums, String region) {
        Pattern skuPattern = Pattern.compile("(\\d+\\.?\\d?R\\d{3})$");
        Map<String, Sheet> sheets = new HashMap<>();

        for (Orders order : ordersMap.values()) {
            boolean matchesRegion = order.getSku().contains(region);
            if (matchesRegion) {
                if (order.getCancelledAt() != null && !order.getCancelledAt().isEmpty()) {
                    writeOrderToSheet(order, workbook.getSheet("Cancelled Orders"), rowNums, "Cancelled Orders", order.getQuantity(), order.getOrderNo());
                } else {
                    Matcher matcher = skuPattern.matcher(order.getSku());
                    if (matcher.find()) {
                        String sku = matcher.group(1);
                        Sheet sheet = sheets.computeIfAbsent(sku, key -> createSheet(workbook, key, TextProcessor.CUSTOM_HEADERS, rowNums));

                        int quantity = order.getQuantity();
                        if (quantity <= 2) {
                            writeOrderToSheet(order, sheet, rowNums, sku, quantity, order.getOrderNo());
                        } else {
                            splitOrderAndWriteToSheet(order, sheet, rowNums, sku, quantity);
                        }
                    }
                }
            }
        }
    }

    private void writeDataToNewSheets(Map<String, Orders> ordersMap, Workbook workbook, String region) {
        Pattern skuPattern = Pattern.compile("(\\d+\\.?\\d?R\\d{3})$");
        Map<String, Sheet> sheets = new HashMap<>();
        Map<String, Integer> rowNums = new HashMap<>();
        Map<String, LinkedHashMap<String, Integer>> sheetEmailCountMaps = new HashMap<>();
        Map<String, List<Orders>> ordersByEmailMap = new HashMap<>();

        // Initialize Unfulfilled Orders sheet
        Sheet unfulfilledSheet = workbook.getSheet("Unfulfilled Orders");
        if (unfulfilledSheet == null) {
            unfulfilledSheet = workbook.createSheet("Unfulfilled Orders");
            Row headerRow = unfulfilledSheet.createRow(0);
            for (int i = 0; i < TextProcessor.NEW_HEADERS.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(TextProcessor.NEW_HEADERS.get(i));
            }
            rowNums.put("UnfulfilledSheet", 1);
            sheetEmailCountMaps.put("Unfulfilled Orders", new LinkedHashMap<>());
        }

        for (Orders order : ordersMap.values()) {
            boolean matchesRegion = order.getSku().contains(region);
            if (matchesRegion && (order.getCancelledAt() == null || order.getCancelledAt().isEmpty())) {
                Matcher matcher = skuPattern.matcher(order.getSku());
                if (matcher.find()) {
                    String sku = matcher.group(1);
                    String sheetName = sku + "_NEW";
                    sheets.computeIfAbsent(sheetName, key -> createSheet(workbook, key, TextProcessor.NEW_HEADERS, rowNums));
                    sheetEmailCountMaps.computeIfAbsent(sheetName, _ -> new LinkedHashMap<>());
                    ordersByEmailMap.computeIfAbsent(sheetName, _ -> new ArrayList<>()).add(order);
                    // Renew email count
                    incrementEmailCount(sheetEmailCountMaps.get(sheetName), order.getEmail());
                }
            }
        }

        // Sort the email count maps
        for (Map.Entry<String, LinkedHashMap<String, Integer>> entry : sheetEmailCountMaps.entrySet()) {
            LinkedHashMap<String, Integer> sortedEmailCountMap = sortByValue(entry.getValue());
            sheetEmailCountMaps.put(entry.getKey(), sortedEmailCountMap);
        }

        for (Map.Entry<String, List<Orders>> entry : ordersByEmailMap.entrySet()) {
            String sheetName = entry.getKey();
            List<Orders> ordersList = entry.getValue();
            Sheet sheet = sheets.get(sheetName);
            final LinkedHashMap<String, Integer> emailCountMap = sheetEmailCountMaps.get(sheetName);

            // Sort orders by email count in descending order
            ordersList.sort((o1, o2) -> emailCountMap.get(o2.getEmail()) - emailCountMap.get(o1.getEmail()));

            for (Orders order : ordersList) {
                boolean matchesRegion = order.getSku().contains(region);
                if (matchesRegion && (order.getCancelledAt() == null || order.getCancelledAt().isEmpty())) {
                    Matcher matcher = skuPattern.matcher(order.getSku());
                    if (matcher.find()) {
                        String sku = matcher.group(1);
                        sheet = sheets.get(sku + "_NEW");

                        int quantity = order.getQuantity();
                        int emailCount = emailCountMap.get(order.getEmail());

                        // Check for unfulfilled condition and write to Unfulfilled Orders sheet
                        if (order.getFulfillmentStatus().equalsIgnoreCase("unfulfilled") || order.getFulfillmentAt().isEmpty()) {
                            writeOrderToNewSheet(order, unfulfilledSheet, rowNums, "UnfulfilledSheet", order.getQuantity(), order.getOrderNo(), emailCount);
                        }

                        if (quantity <= 2) {
                            writeOrderToNewSheet(order, sheet, rowNums, sheetName, quantity, order.getOrderNo(), emailCount);
                        } else {
                            splitOrderAndWriteToNewSheet(order, sheet, rowNums, sheetName, quantity, emailCount);
                        }
                    }
                }
            }
        }
    }

    private void incrementEmailCount(Map<String, Integer> emailCountMap, String email) {
        emailCountMap.put(email, emailCountMap.getOrDefault(email, 0) + 1);
    }

    private LinkedHashMap<String, Integer> sortByValue(Map<String, Integer> unsortedMap) {
        List<Map.Entry<String, Integer>> list = new ArrayList<>(unsortedMap.entrySet());
        list.sort((entry1, entry2) -> entry2.getValue().compareTo(entry1.getValue()));

        LinkedHashMap<String, Integer> sortedMap = new LinkedHashMap<>();
        for (Map.Entry<String, Integer> entry : list) {
            sortedMap.put(entry.getKey(), entry.getValue());
        }
        return sortedMap;
    }

    private Sheet createSheet(Workbook workbook, String sheetName, List<String> headers, Map<String, Integer> rowNums) {
        Sheet sheet = workbook.createSheet(sheetName);
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
        }
        rowNums.put(sheetName, 1);
        return sheet;
    }

    private void writeOrderToSheet(Orders order, Sheet sheet, Map<String, Integer> rowNums, String sheetName, int quantity, String orderNo) {
        int rowNum = rowNums.getOrDefault(sheetName, 1);
        Row row = sheet.createRow(rowNum);
        for (int i = 0; i < TextProcessor.CUSTOM_HEADERS.size(); i++) {
            Cell cell = row.createCell(i);
            switch (i) {
                case 0:
                    cell.setCellValue(orderNo);
                    break;
                case 1:
                    cell.setCellValue(order.getEmail());
                    break;
                case 2:
                    cell.setCellValue(order.getSku());
                    break;
                case 3:
                    cell.setCellValue(order.getDisCountCode());
                    break;
                case 4:
                    cell.setCellValue(order.getProduct());
                    break;
                case 5:
                    cell.setCellValue(quantity);
                    break;
                case 6:
                    cell.setCellValue(order.getTotalPrice());
                    break;
                case 7:
                    String processedName = processShippingName(order.getName(), cell);
                    cell.setCellValue(processedName);
                    break;
                case 8:
                    cell.setCellValue(order.getAddress());
                    if (isEntirelyEnglish(order.getAddress())) {
                        highlightCell(cell);
                    }
                    break;
                case 9:
                    cell.setCellValue(order.getCountry());
                    break;
                case 10:
                    String processedPhone = processPhone(order.getPhone().trim(), cell);
                    cell.setCellValue(processedPhone);
                    break;
                case 11:
                    cell.setCellValue(String.join(", ", order.getNotes()));
                    break;
                case 12:
                    if (order.isMerged()) {
                        cell.setCellValue(order.getFulfillmentNote());
                    }
                    break;
                default:
                    break;
            }
        }
        rowNums.put(sheetName, rowNum + 1);
    }

    private void writeOrderToNewSheet(Orders order, Sheet sheet, Map<String, Integer> rowNums, String sheetName, int quantity, String orderNo, int emailCount) {
        int rowNum = rowNums.getOrDefault(sheetName, 1);
        Row row = sheet.createRow(rowNum);
        for (int i = 0; i < NEW_HEADERS.size(); i++) {
            Cell cell = row.createCell(i);
            switch (i) {
                case 0:
                    cell.setCellValue(""); // 出貨日
                    break;
                case 1:
                    cell.setCellValue(""); // 指定配達日
                    break;
                case 2:
                    cell.setCellValue(orderNo); // 訂單編號
                    break;
                case 3:
                    cell.setCellValue(order.getTotalPrice() / order.getQuantity() * quantity); // 代收金額
                    break;
                case 4:
                    cell.setCellValue(order.getProduct() + "* " + quantity + "箱"); // 品項
                    break;
                case 5:
                    cell.setCellValue(String.join(", ", order.getNotes())); // 備註
                    break;
                case 6:
                    String processedName = processShippingName(order.getName(), cell); // 收件人
                    cell.setCellValue(processedName);
                    checkAndHighlightEmptyCell(cell);
                    break;
                case 7:
                    String processedPhone = processPhone(order.getPhone().trim(), cell);
                    cell.setCellValue(processedPhone); // 手機(收)
                    checkAndHighlightEmptyCell(cell);
                    break;
                case 8:
                    cell.setCellValue(""); // 電話(收)
                    break;
                case 9:
                    cell.setCellValue(order.getAddress()); // 地址(收)
                    if (isEntirelyEnglish(order.getAddress())) {
                        highlightCell(cell);
                    }
                    checkAndHighlightEmptyCell(cell);
                    break;
                case 10:
                    cell.setCellValue(""); // 契客代號
                    break;
                case 11:
                    cell.setCellValue("2"); // 溫層
                    break;
                case 12:
                    cell.setCellValue("1"); // 規格
                    break;
                case 13:
                    cell.setCellValue("1"); // 配送時間帶
                    break;
                case 14:
                    cell.setCellValue(""); // 會員編號
                    break;
                case 15:
                    cell.setCellValue("iWoW"); // 寄件人
                    break;
                case 16:
                    cell.setCellValue(""); // 寄件人地址
                    break;
                case 17:
                    cell.setCellValue(""); // 寄件人電話
                    break;
                case 18:
                    cell.setCellValue(""); // 複數件
                    break;
                case 19:
                    cell.setCellValue(quantity);
                    break;
                case 20:
                    cell.setCellValue(order.getEmail());
                    break;
                case 21:
                    cell.setCellValue(emailCount);
                    break;
                case 22:
                    if (order.isMerged()) {
                        cell.setCellValue(order.getFulfillmentNote());
                    }
                    break;
                default:
                    break;
            }
        }
        rowNums.put(sheetName, rowNum + 1);
    }

    private void checkAndHighlightEmptyCell(Cell cell) {
        if (cell.getCellType() == CellType.BLANK || (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty())) {
            highlightCell(cell);
        }
    }

    private void splitOrderAndWriteToSheet(Orders order, Sheet sheet, Map<String, Integer> rowNums, String sku, int quantity) {
        int splitCount = 1;
        double unitPrice = order.getTotalPrice() / order.getQuantity();

        while (quantity > 0) {
            int currentSplitQty = Math.min(quantity, 2);
            quantity -= currentSplitQty;

            String splitOrderNo = order.getOrderNo() + "-" + splitCount;
            splitCount++;
//            double splitTotalPrice = unitPrice * currentSplitQty;
//            order.setTotalPrice(splitTotalPrice);

            writeOrderToSheet(order, sheet, rowNums, sku, currentSplitQty, splitOrderNo);
        }
    }

    private void splitOrderAndWriteToNewSheet(Orders order, Sheet sheet, Map<String, Integer> rowNums, String sku, int quantity, int emailCount) {
        int splitCount = 1;
        double unitPrice = order.getTotalPrice() / order.getQuantity();

        while (quantity > 0) {
            int currentSplitQty = Math.min(quantity, 2);
            quantity -= currentSplitQty;

            String splitOrderNo = order.getOrderNo() + "-" + splitCount;
            splitCount++;
//            double splitTotalPrice = unitPrice * currentSplitQty;
//            System.out.println("SplitTotalPrice: " + splitTotalPrice);
//            order.setTotalPrice(splitTotalPrice);
            writeOrderToNewSheet(order, sheet, rowNums, sku, currentSplitQty, splitOrderNo, emailCount);
        }
    }

    //  Process Customer Name
    private String processShippingName(String shippingName, Cell cell) {
        shippingName = shippingName.trim();
        boolean containsEnglish = containsEnglishCharacter(shippingName);

        int lastSpaceIndex = shippingName.lastIndexOf(' ');
        if (lastSpaceIndex > 0) {
            String lastName = shippingName.substring(lastSpaceIndex + 1);
            String firstName = shippingName.substring(0, lastSpaceIndex);
            if (containsEnglish) {
                highlightCell(cell);
                return firstName + " " + lastName; // For English names
            } else {
                return lastName + firstName; // For non-English names
            }
        }
        return shippingName;
    }

    private boolean containsEnglishCharacter(String str) {
        for (char c : str.toCharArray()) {
            if (isEnglishCharacter(c)) {
                return true;
            }
        }
        return false;
    }

    private boolean isEnglishCharacter(char c) {
        return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
    }

    private void highlightCell(Cell cell) {
        Workbook workbook = cell.getSheet().getWorkbook();
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(style);
    }

    // Process All Address not contained any chinese correctors
    private boolean isEntirelyEnglish(String str) {
        for (char c : str.toCharArray()) {
            if (!isEnglishCharacter(c) && !isWhitespaceOrPunctuation(c)) {
                return false;
            }
        }
        return true;
    }

    private boolean isWhitespaceOrPunctuation(char c) {
        return Character.isWhitespace(c) || Character.isDigit(c) || isPunctuation(c);
    }

    private boolean isPunctuation(char c) {
        return "!@#$%^&*()_+-=[]{}|;':\",.<>?/".indexOf(c) >= 0;
    }

    //  Process Customer Phone
    private String processPhone(String rawPhone, Cell cell) {
        String cleanedPhone = rawPhone.trim().replace(" ", "").replace("'", "");
        if (cleanedPhone.startsWith("+886")) {
            cleanedPhone = cleanedPhone.replace("+886", "0").trim();
        } else if (cleanedPhone.startsWith("886")) {
            cleanedPhone = cleanedPhone.replace("886", "0").trim();
        } else if (cleanedPhone.startsWith("+1")) {
            cleanedPhone = cleanedPhone.replace("+1", "").trim();

            // Check if the phone number matches the criteria then highlight
            highlightCell(cell);
        }

        return cleanedPhone;
    }

    private void saveWorkbook(Workbook workbook, String outputFileName) {
        try (FileOutputStream fileOut = new FileOutputStream(outputFileName)) {
            workbook.write(fileOut);
        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, "saveWorkbook error", e);
        }
    }
}
