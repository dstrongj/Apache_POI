import java.io.File;
import java.io.FileOutputStream;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDatabase {
   public static void main(String[] args) throws Exception {
      Class.forName("oracle.jdbc.driver.OracleDriver");
      Connection connect = DriverManager.getConnection("jdbc:oracle:thin:USERNAME/PASSWORD@DB_CONNECTION_STRING");
      
      Statement statement = connect.createStatement();
      ResultSet resultSet = statement.executeQuery("SQL SELECT STATEMENT HERE");
      XSSFWorkbook workbook = new XSSFWorkbook(); 
      XSSFSheet spreadsheet = workbook.createSheet("Sample_Workbook");
      
      XSSFRow row = spreadsheet.createRow(1);
      XSSFCell cell;
      cell = row.createCell(1);
      cell.setCellValue("Column_1");
      cell = row.createCell(2);
      cell.setCellValue("Column_2");
      cell = row.createCell(3);
      cell.setCellValue("Column_3");
      cell = row.createCell(4);
      cell.setCellValue("Column_4");
      cell = row.createCell(5);
      cell.setCellValue("Column_5");
      cell = row.createCell(6);
      cell.setCellValue("Column_6");
      int i = 2;

      while(resultSet.next()) {
         row = spreadsheet.createRow(i);
         cell = row.createCell(1);
         cell.setCellValue(resultSet.getString("Column_1"));
         cell = row.createCell(2);
         cell.setCellValue(resultSet.getString("Column_2"));
         cell = row.createCell(3);
         cell.setCellValue(resultSet.getString("Column_3"));
         cell = row.createCell(4);
         cell.setCellValue(resultSet.getString("Column_4"));
         cell = row.createCell(5);
         cell.setCellValue(resultSet.getString("Column_5"));
         cell = row.createCell(6);
         cell.setCellValue(resultSet.getString("Column_6"));
         i++;
      }

      FileOutputStream out = new FileOutputStream(new File("Sample_File.csv"));
      workbook.write(out);
      out.close();
      System.out.println("Sample_File.csv written successfully");
   }
}