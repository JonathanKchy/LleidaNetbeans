/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.lleida_sodig;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.net.ssl.HttpsURLConnection;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Andrés Aymacaña
 */
public class SodigLleida {
     public static void main(String[] args) throws MalformedURLException, IOException {
        System.out.println("Hola mundo");
          int contador=0;
        
         CrearExcel();
    }
    
    //crear excel
    
    public static void CrearExcel() {
    
        Workbook book=new XSSFWorkbook();
        Sheet sheet=book.createSheet("Hola Java");
        
         try {
             FileOutputStream fileout=new FileOutputStream("Excel.xlsx");
            try {
                book.write(fileout);
            } catch (IOException ex) {
                Logger.getLogger(SodigLleida.class.getName()).log(Level.SEVERE, null, ex);
            }
            try {
                fileout.close();
            } catch (IOException ex) {
                Logger.getLogger(SodigLleida.class.getName()).log(Level.SEVERE, null, ex);
            }
             
         } catch (FileNotFoundException ex) {
             Logger.getLogger(SodigLleida.class.getName()).log(Level.SEVERE, null, ex);
         }
    }

}
