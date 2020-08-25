/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Test;

import Core.ExcelToDb;
import java.io.IOException;



/**
 *
 * @author DATA PC
 */
public class PoiTest {
   
    public static void main(String[]args) throws IOException{
        String filename = "C:\\\\Users\\\\DATA PC\\\\Documents\\\\NetBeansProjects\\\\ExcelPoi\\\\src\\\\java\\\\Document\\\\SampleData.xlsx";
        ExcelToDb.ExcelToDb(filename);
    }
   
}
