/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Core;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author DATA PC
 */
public class DbToExcel {

    public static void dbtoex(String filename ) throws ClassNotFoundException, SQLException, FileNotFoundException, IOException {

        //Connecting to the database
        Class.forName("org.apache.derby.jdbc.ClientDriver");
        Connection con = DriverManager.getConnection("jdbc:derby://localhost:1527/exceldata", "abc", "123");

        Statement statement = con.createStatement();
        ResultSet resultSet = statement.executeQuery("select * from excel");

        File file = new File(filename);
        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        //Creating a Spread Sheet
        XSSFSheet spreadsheet = workbook.createSheet("SalesOrder3");
        XSSFRow row = spreadsheet.createRow(0);
        XSSFCell cell;

        cell = row.createCell(0);
        cell.setCellValue("REGION");

        cell = row.createCell(1);
        cell.setCellValue("REP");

        cell = row.createCell(2);
        cell.setCellValue("ITEM");

        cell = row.createCell(3);
        cell.setCellValue("UNIT");

        cell = row.createCell(4);
        cell.setCellValue("UNIT COST");

        cell = row.createCell(5);
        cell.setCellValue("TOTAL");

        int i = 1;
        while (resultSet.next()) {
            row = spreadsheet.createRow(i);
            cell = row.createCell(0);
            cell.setCellValue(resultSet.getString("region"));

            cell = row.createCell(1);
            cell.setCellValue(resultSet.getString("rep"));

            cell = row.createCell(2);
            cell.setCellValue(resultSet.getString("Item"));

            cell = row.createCell(3);
            cell.setCellValue(resultSet.getString("Unit"));

            cell = row.createCell(4);
            cell.setCellValue(resultSet.getString("UnitCost"));

            cell = row.createCell(5);
            cell.setCellValue(resultSet.getString("Total"));
            System.out.println("Importing row " + i);
            i++;
        }

        try (FileOutputStream out = new FileOutputStream(
                new File(filename))) {
            workbook.write(out);
        }

        System.out.println("MSG: Write To Excel File Successful!");
    }
}
