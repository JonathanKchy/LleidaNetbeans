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
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
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
        // Creamos el libro de trabajo de Excel formato OOXML
        Workbook book=new XSSFWorkbook();
         // La hoja donde pondremos los datos
        Sheet sheet=book.createSheet("Hola Java");
        //creamos una fila
        Row row=sheet.createRow(contador);
        row.createCell(0).setCellValue("Id");
        row.createCell(1).setCellValue("Fecha Lleida");
        
        
        
        try {
         
        // URL url= new URL("https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_id=100274269");
       URL url= new URL("https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_date_min=20220501070000&mail_date_max=20220601070000");
        HttpsURLConnection conn=(HttpsURLConnection)url.openConnection();
        conn.setRequestMethod("GET");
        conn.connect();
           
        int responseCod= conn.getResponseCode();
            if (responseCod!=200) {
                throw new RuntimeException("ocurrio un error: "+responseCod);
            }else  {
                StringBuilder informationString=new StringBuilder();
                Scanner scanner=new Scanner(url.openStream());
                while (scanner.hasNext()) {                    
                    String linea=scanner.nextLine();
                    
                    //informationString.append("\n");
                    if (linea.contains("<mail_id>")) {
                        int tamano=linea.length();
                        int fin=tamano-10;
                        contador ++;
                        row=sheet.createRow(contador);
                        row.createCell(0).setCellValue(linea.substring(9, fin));
                        
                        //System.out.println(linea.length());
                        //System.out.println(linea);
                        //System.out.println("mail_id: "+linea.substring(9, fin));
                        informationString.append("mail_id: "+linea.substring(9, fin));
                        informationString.append("\n");
                    }else if(linea.contains("<mail_date>")) {
                        int tamano=linea.length();
                        int fin=tamano-12;
                        row.createCell(1).setCellValue(linea.substring(11, fin));
                        //System.out.println(linea.length());
                        //System.out.println(linea);
                        //System.out.println("mail_id: "+linea.substring(9, fin));
                        informationString.append("mail_date: "+linea.substring(11, fin));
                        informationString.append("\n");
                    }else if(linea.contains("</pdf_row>")) {
                        int tamano=linea.length();
                        int fin=tamano-10;
                        //System.out.println(linea.length());
                        //System.out.println(linea);
                        //System.out.println("mail_id: "+linea.substring(9, fin));
                       // informationString.append("gstatus: "+linea.substring(9, fin));
                        informationString.append("\n");
                    }
                                       
                }
                scanner.close();
                System.out.println(informationString);
                System.out.println(contador);
            }
           // System.out.println("Elemento raiz:"+documentoXML.getDocumentElement().getNodeName());
            
           
        } catch (Exception e) {
            System.out.println("error: ");
        }
    
    
        
        
         CrearExcel(book);
    }
    
    //crear excel
    
    public static void CrearExcel(Workbook book) {
    
        // Creamos el estilo paga las celdas del encabezado
        CellStyle style = book.createCellStyle();
        // Indicamos que tendra un fondo azul aqua
        // con patron solido del color indicado
        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
         try {
            // Creamos el flujo de salida de datos,
            // apuntando al archivo donde queremos 
            // almacenar el libro de Excel
             FileOutputStream fileout=new FileOutputStream("Excel.xlsx");
            try {
                // Almacenamos el libro de 
            // Excel via ese 
            // flujo de datos
                book.write(fileout);
            } catch (IOException ex) {
                Logger.getLogger(SodigLleida.class.getName()).log(Level.SEVERE, null, ex);
            }
            try {
                // Cerramos el libro para concluir operaciones
                fileout.close();
               //  LOGGER.log(Level.INFO, "Archivo creado existosamente en {0}", fil.getAbsolutePath());
            } catch (IOException ex) {
                Logger.getLogger(SodigLleida.class.getName()).log(Level.SEVERE, null, ex);
            }
             
         } catch (FileNotFoundException ex) {
             Logger.getLogger(SodigLleida.class.getName()).log(Level.SEVERE, null, ex);
         }
    }

}
