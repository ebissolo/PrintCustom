/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package printcustom;

import java.io.File;
import java.io.IOException;
import jxl.Cell;
import jxl.DateCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 *
 * @author ebissolo
 */
public class GetCells {
    
    Cell temp;
    DateCell date;
    String[] Dati = new String[3];
    String[] Cliente = new String[5]; 
    String[] Quantità = new String[14];
    String[] Descrizione = new String[14];
    String[] PrezzoUnit = new String[14];
    String IVA = "";
    int[] pos_quantità = new int[14];
    int[] pos_descrizione = new int[14];
    int[] pos_prezzounit = new int[14];
    Sheet sheet;
    Workbook workbook;
    
    public GetCells() { }
    
    public void openWorkbook(String nameWorkbook) {
        try {        
            workbook = Workbook.getWorkbook(new File(nameWorkbook));
            sheet = workbook.getSheet(0);
        } catch (BiffException | IOException e) {}
    }
   
    public Workbook getWorkBook(){
        return workbook;
    }
    
    public void leggiDati(){
            temp = sheet.getCell(0, 10);
            Dati[0] = temp.getContents();
            temp = sheet.getCell(2, 10);
            Dati[1] = temp.getContents();
            temp = sheet.getCell(4, 10);
            Dati[2] = temp.getContents();
    }
    
    public void infoCliente(){
            for(int i = 14;i < 19;i ++){
                temp = sheet.getCell(0, i);
                Cliente[i-14] = temp.getContents();
            }
    }
    
    public int infoQuantità(){
            int index = 0;
            for(int i=23;i < 37;i ++){
                if((temp = sheet.getCell(0, i)) != null){
                    Quantità[index] = temp.getContents();
                    pos_quantità[index++] = i;
                 }
            }
            //serve per sapere quanto far ciclare il for per la stampa sullo sheet vuoto
            return index;
    }
    
    public int infoDescrizione(){
            int index = 0;
            for(int i=23;i < 37;i ++){
                if((temp = sheet.getCell(1, i)) != null){
                    Descrizione[index] = temp.getContents();
                    pos_descrizione[index++] = i;
                 }
            }
            //serve per sapere quanto far ciclare il for per la stampa sullo sheet vuoto
            return index;
    }
    
    public int infoPrezzoUnit(){
            int index = 0;
            for(int i=23;i < 37;i ++){
                if((temp = sheet.getCell(4, i)) != null){
                    PrezzoUnit[index] = temp.getContents();
                    pos_prezzounit[index++] = i;
                 }
            }
            //serve per sapere quanto far ciclare il for per la stampa sullo sheet vuoto
            return index;
    }
    
    public String infoIVA(){
            temp = sheet.getCell(5, 38);
            if(temp == null) {
                return "";
            }
            else{
                IVA = temp.getContents();
                //String subIVA = IVA.substring(0, 2);
                return IVA;
            }
    }
    
    public void stampaDati(){
            System.out.println(Dati[0] + " " + Dati[1] + " " + Dati[2]);
            for(int i = 0;i < 5;i ++){
                System.out.println(Cliente[i]);
            }
            for(int i = 0;i < 14;i ++){
                System.out.println(Quantità[i] + "-->" + pos_quantità[i]);
            }
            for(int i = 0;i < 14;i ++){
                System.out.println(Descrizione[i] + "-->" + pos_descrizione[i]);
            }       
            for(int i = 0;i < 14;i ++){
                System.out.println(PrezzoUnit[i] + "-->" + pos_prezzounit[i]);
            }
            if(this.infoIVA() != null)
                System.out.println("IVA value = " + this.infoIVA());
            else
                System.out.println("IVA value missing!");
    }
    
    public void closeWorkbook(){
        workbook.close();
    }
}
