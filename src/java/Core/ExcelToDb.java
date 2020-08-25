/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Core;

/**
 *
 * @author DATA PC
 */
import java.io.*;
import java.sql.*;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToDb {

    public static void ExcelToDb(String filename) {

        try {
            Class.forName("org.apache.derby.jdbc.ClientDriver");
            try (Connection con = DriverManager.getConnection("jdbc:derby://localhost:1527/exceldata", "abc", "123")) {
                con.setAutoCommit(false);
                File file = new File(filename);
                FileInputStream fis = new FileInputStream(file);
                XSSFWorkbook wb = new XSSFWorkbook(fis);
                XSSFSheet sheet = wb.getSheetAt(0);

                Row row;
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    row = sheet.getRow(i);

                    String region = row.getCell(0).getStringCellValue();
                    String rep = row.getCell(1).getStringCellValue();
                    String item = row.getCell(2).getStringCellValue();
                    double unit = row.getCell(3).getNumericCellValue();
                    double unitcost = row.getCell(4).getNumericCellValue();
                    double total = row.getCell(5).getNumericCellValue();

                    PreparedStatement ps = con.prepareStatement("insert into excel (region,rep,item,unit,unitcost,total) values (?,?,?,?,?,?)");

                    ps.setString(1, region);
                    ps.setString(2, rep);
                    ps.setString(3, item);
                    ps.setDouble(4, unit);
                    ps.setDouble(5, unitcost);
                    ps.setDouble(6, total);
                    ps.executeUpdate();

                    System.out.println("Exporting row " + i);
                }
                con.commit();
                con.close();
            }

            System.out.println("MSG: Exporting Excel Data to Mysql Table Successful");
        } catch (ClassNotFoundException | SQLException | NullPointerException | IOException e) {
            System.out.println(e);
        }

    }
}
