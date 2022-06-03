
package sodiglleida;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.net.ssl.HttpsURLConnection;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;





public class SodigLleida {
 
    
   
    public static void main(String[] args) throws MalformedURLException, IOException {
        System.out.println("Hola mundo");
          int contador=0;
        CrearExcel();
        /*try {
         
        //  URL url= new URL("https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_id=100274269");
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
                        //System.out.println(linea.length());
                        //System.out.println(linea);
                        //System.out.println("mail_id: "+linea.substring(9, fin));
                        informationString.append("mail_id: "+linea.substring(9, fin));
                        informationString.append("\n");
                    }else if(linea.contains("<mail_date>")) {
                        int tamano=linea.length();
                        int fin=tamano-12;
                        //System.out.println(linea.length());
                        //System.out.println(linea);
                        //System.out.println("mail_id: "+linea.substring(9, fin));
                        informationString.append("mail_date: "+linea.substring(11, fin));
                        informationString.append("\n");
                    }else if(linea.contains("<mail_from>")) {
                        int tamano=linea.length();
                        int fin=tamano-12;
                        //System.out.println(linea.length());
                        //System.out.println(linea);
                        //System.out.println("mail_id: "+linea.substring(9, fin));
                        informationString.append("mail_from: "+linea.substring(11, fin));
                        informationString.append("\n");
                    }else if(linea.contains("<mail_to>")) {
                        int tamano=linea.length();
                        int fin=tamano-10;
                        //System.out.println(linea.length());
                        //System.out.println(linea);
                        //System.out.println("mail_id: "+linea.substring(9, fin));
                        informationString.append("mail_to: "+linea.substring(9, fin));
                        informationString.append("\n");
                    }else if(linea.contains("<mail_type>")) {
                        int tamano=linea.length();
                        int fin=tamano-12;
                        //System.out.println(linea.length());
                        //System.out.println(linea);
                        //System.out.println("mail_id: "+linea.substring(9, fin));
                        informationString.append("mail_type: "+linea.substring(11, fin));
                        informationString.append("\n");
                    }else if(linea.contains("<mail_subj>")) {
                        int tamano=linea.length();
                        int fin=tamano-12;
                        //System.out.println(linea.length());
                        //System.out.println(linea);
                        //System.out.println("mail_id: "+linea.substring(9, fin));
                        informationString.append("mail_subj: "+linea.substring(11, fin));
                        informationString.append("\n");
                    }else if(linea.contains("<gstatus>")) {
                        int tamano=linea.length();
                        int fin=tamano-10;
                        //System.out.println(linea.length());
                        //System.out.println(linea);
                        //System.out.println("mail_id: "+linea.substring(9, fin));
                        informationString.append("gstatus: "+linea.substring(9, fin));
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
        }*/
    }
    
    //crear excel
   public static void CrearExcel() {
    
        Workbook book=new HSSFWorkbook();
        Sheet sheet=book.createSheet("Hola Java");
        
         try {
             FileOutputStream fileout=new FileOutputStream("Excel.xls");
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
