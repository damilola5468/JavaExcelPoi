/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Test;

import Core.DbToExcel;
import Core.ExcelToDb;
import java.io.IOException;
import java.sql.SQLException;

/**
 *
 * @author DATA PC
 */
public class PoiTest {

    public static void main(String[] args) throws IOException, ClassNotFoundException, SQLException {
//        String filename = "C:\\\\Users\\\\DATA PC\\\\Documents\\\\NetBeansProjects\\\\ExcelPoi\\\\src\\\\java\\\\Document\\\\SampleData.xlsx";
//        ExcelToDb.ExcelToDb(filename);
          DbToExcel.dbtoex();
    }

}
