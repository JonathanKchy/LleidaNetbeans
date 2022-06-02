
package sodiglleida;

import java.io.File;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Scanner;
import javax.net.ssl.HttpsURLConnection;



public class SodigLleida {
 
    
   
    public static void main(String[] args) throws MalformedURLException {
        System.out.println("Hola mundo");
          
        try {
         
           URL url= new URL("https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_id=100274269");
          //URL url= new URL("https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_date_min=20220602000000");
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
                    informationString.append(linea);
                    informationString.append("\n");
                    if (linea.contains("<mail_id>")) {
                        System.out.println(linea.length());
                        System.out.println(linea);
                    }
                                       
                }
                scanner.close();
                //System.out.println(informationString);
            }
           // System.out.println("Elemento raiz:"+documentoXML.getDocumentElement().getNodeName());
            
           
        } catch (Exception e) {
            System.out.println("error: ");
        }
                
    }
    
}
