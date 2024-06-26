import java.io.*;
import java.nio.file.Paths;
import java.sql.*;
import java.util.*;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DatabaseToExcel {
    private static final Logger logger = LogManager.getLogger(DatabaseToExcel.class);
    private static final String URL = "jdbc:oracle:thin:@//vm-oracle19x.eiskonzept.ag:1521/seatdb1x.eiskonzept.ag";
    private static final String USER = "MARIJA";
    private static final String PASSWORD = "test";

    public static void main(String[] args) {
        try (Scanner scanner = new Scanner(System.in)) {
            Connection conn = null;
            try {
                // Initialize logger
                logger.info("Application started.");

                // Establish connection
                conn = DriverManager.getConnection(URL, USER, PASSWORD);
                logger.info("Connected to the database!");

                // Prompt user for report name
                System.out.print("Enter the report name (e.g., report1): ");
                String reportName = scanner.nextLine().trim();

                // Get SQL query from configuration
                String query = getQueryFromConfig(reportName);
                if (query == null) {
                    logger.warn("Report '{}' not found in the configuration.", reportName);
                    System.out.println("Report not found in the configuration.");
                    return;
                }

                // Prompt user for file name to save the Excel file
                System.out.print("Enter file name to save the Excel file (e.g., result): ");
                String fileName = scanner.nextLine().trim();

                // Append .xlsx extension if not already present
                if (!fileName.endsWith(".xlsx")) {
                    fileName += ".xlsx";
                }

                // Create Excel file in the current working directory
                String filePath = Paths.get("").toAbsolutePath().toString() + File.separator + fileName;

                // Write data to Excel file
                try (Statement stmt = conn.createStatement();
                     ResultSet rs = stmt.executeQuery(query)) {
                    writeExcel(rs, filePath);
                    logger.info("Export successful. File saved as: {}", filePath);
                    System.out.println("Export successful. File saved as: " + filePath);
                } catch (SQLException e) {
                    logger.error("Query execution problem: {}", e.getMessage(), e);
                    System.err.println("Query execution problem: " + e.getMessage());
                }
            } catch (SQLException e) {
                logger.error("Connection failed: {}", e.getMessage(), e);
                System.err.println("Connection failed: " + e.getMessage());
            } finally {
                // Close connection
                if (conn != null) {
                    try {
                        conn.close();
                    } catch (SQLException e) {
                        logger.error("Error closing connection: {}", e.getMessage(), e);
                        e.printStackTrace();
                    }
                }
            }
        } catch (Exception e) {
            logger.error("Unexpected error: {}", e.getMessage(), e);
            e.printStackTrace();
        } finally {
            // Clean up resources if needed
        }
    }

    // Method to get SQL query from configuration file
    private static String getQueryFromConfig(String reportName) {
        Properties prop = new Properties();
        try (InputStream input = new FileInputStream("reportConfig.properties")) {
            prop.load(input);
            return prop.getProperty(reportName);
        } catch (IOException ex) {
            logger.error("Error reading report configuration file: {}", ex.getMessage(), ex);
            ex.printStackTrace();
        }
        return null;
    }

    // Method to write ResultSet to Excel file with specified file name
    private static void writeExcel(ResultSet rs, String filePath) {
        List<String> configColumns = readConfig();
        File file = new File(filePath);
        XSSFWorkbook workbook;
        Sheet sheet;
        Map<String, Integer> columnMapping = new HashMap<>();

        try {
            if (file.exists()) {
                // Read existing file
                try (FileInputStream fis = new FileInputStream(file)) {
                    workbook = new XSSFWorkbook(fis);
                    sheet = workbook.getSheetAt(0);

                    // Read existing headers
                    Row headerRow = sheet.getRow(0);
                    if (headerRow != null) {
                        for (Cell cell : headerRow) {
                            columnMapping.put(cell.getStringCellValue(), cell.getColumnIndex());
                        }
                    }
                }
            } else {
                // Create new file
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet("Data");
            }

            // Create header row if it doesn't exist
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                headerRow = sheet.createRow(0);
            }

            // Create or update header row using ResultSet metadata
            int columnCount = rs.getMetaData().getColumnCount();
            for (String columnName : configColumns) {
                for (int i = 1; i <= columnCount; i++) {
                    if (columnName.equalsIgnoreCase(rs.getMetaData().getColumnLabel(i))) {
                        if (!columnMapping.containsKey(columnName)) {
                            int newColumnIndex = headerRow.getLastCellNum();
                            if (newColumnIndex < 0) {
                                newColumnIndex = 0;
                            }
                            headerRow.createCell(newColumnIndex).setCellValue(columnName);
                            columnMapping.put(columnName, newColumnIndex);
                        }
                    }
                }
            }

            // Create data rows
            int rowNum = sheet.getLastRowNum() + 1;
            while (rs.next()) {
                Row row = sheet.createRow(rowNum++);
                for (String columnName : configColumns) {
                    for (int i = 1; i <= columnCount; i++) {
                        if (columnName.equalsIgnoreCase(rs.getMetaData().getColumnLabel(i))) {
                            int colIndex = columnMapping.get(columnName);
                            row.createCell(colIndex).setCellValue(rs.getString(i));
                        }
                    }
                }
            }

            // Write to file
            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
            }
        } catch (IOException | SQLException e) {
            logger.error("Error writing Excel file: {}", e.getMessage(), e);
            e.printStackTrace();
        }
    }

    // Method to read configuration
    private static List<String> readConfig() {
        List<String> columns = new ArrayList<>();
        try (InputStream input = new FileInputStream("columnConfig.properties")) {
            Properties prop = new Properties();
            prop.load(input);
            String columnsStr = prop.getProperty("columns");
            if (columnsStr != null) {
                columns = Arrays.asList(columnsStr.split(","));
            }
        } catch (IOException ex) {
            logger.error("Error reading column configuration file: {}", ex.getMessage(), ex);
            ex.printStackTrace();
        }
        return columns;
    }
}
