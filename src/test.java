import java.io.*;
import java.nio.file.Paths;
import java.sql.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class test {
    private static final String URL = "jdbc:oracle:thin:@//vm-oracle19x.eiskonzept.ag:1521/seatdb1x.eiskonzept.ag";
    private static final String USER = "MARIJA";
    private static final String PASSWORD = "test";

    public static void main(String[] args) {
        try (Scanner scanner = new Scanner(System.in)) {
            Connection conn = null;
            try {
                // Establish connection
                conn = DriverManager.getConnection(URL, USER, PASSWORD);
                System.out.println("Connected to the database!");

                // Prompt user for report name
                System.out.print("Enter the report name (e.g., report1): ");
                String reportName = scanner.nextLine().trim();

                // Get SQL query from configuration
                String query = getQueryFromConfig(reportName);
                if (query == null) {
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
                String filePath = Paths.get("").toAbsolutePath().toString() + "\\" + fileName;
                try (Statement stmt = conn.createStatement();
                     ResultSet rs = stmt.executeQuery(query)) {
                    // Write data to Excel file
                    writeExcel(rs, filePath);
                    System.out.println("Export successful. File saved as: " + filePath);
                } catch (SQLException e) {
                    System.err.println("Query execution problem: " + e.getMessage());
                }
            } catch (SQLException e) {
                System.err.println("Connection failed: " + e.getMessage());
            } finally {
                // Close connection
                if (conn != null) {
                    try {
                        conn.close();
                    } catch (SQLException e) {
                        e.printStackTrace();
                    }
                }
            }
        }
    }

    // Method to get SQL query from configuration file
    private static String getQueryFromConfig(String reportName) {
        Properties prop = new Properties();
        try (InputStream input = new FileInputStream("reportConfig.properties")) {
            prop.load(input);
            return prop.getProperty(reportName);
        } catch (IOException ex) {
            ex.printStackTrace();
        }
        return null;
    }

    // Method to write ResultSet to Excel file with specified file name
    private static void writeExcel(ResultSet rs, String filePath) {
        File file = new File(filePath);
        XSSFWorkbook workbook;
        Sheet sheet;

        try {
            if (file.exists()) {
                // Read existing file
                try (FileInputStream fis = new FileInputStream(file)) {
                    workbook = new XSSFWorkbook(fis);
                    sheet = workbook.getSheetAt(0);
                }
            } else {
                // Create new file
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet("Data");
            }

            // Create header row using ResultSet metadata
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                headerRow = sheet.createRow(0);
            }
            int columnCount = rs.getMetaData().getColumnCount();
            for (int i = 1; i <= columnCount; i++) {
                String columnName = rs.getMetaData().getColumnLabel(i);
                Cell cell = headerRow.createCell(i - 1);
                cell.setCellValue(columnName);
            }

            // Create data rows
            int rowNum = sheet.getLastRowNum() + 1;
            while (rs.next()) {
                Row row = sheet.createRow(rowNum++);
                for (int i = 1; i <= columnCount; i++) {
                    row.createCell(i - 1).setCellValue(rs.getString(i));
                }
            }

            // Write to file
            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
            }
        } catch (IOException | SQLException e) {
            e.printStackTrace();
        }
    }
}
