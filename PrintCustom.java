/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package printcustom;

import java.io.IOException;

/**
 *
 * @author ebissolo
 */
public class PrintCustom {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {
         
                int index = 0;
                int indice_quantità, indice_descrizione, indice_prezzounit;
                String IVA_value = "";
                        
                //Create ogjects
                GetCells readcells = new GetCells();
                PrintCells printcells = new PrintCells();
                    
                //read data from first file
                readcells.openWorkbook("Ricevuta_campi_pieni.xls");
                readcells.leggiDati();
                readcells.infoCliente();
                readcells.infoQuantità();
                
                //print data to the second file
                printcells.printData(10, 0, readcells.Dati[0]);
                printcells.printData(10, 2, readcells.Dati[1]);
                printcells.printData(10, 4, readcells.Dati[2]);
                
                for(int i = 14;i < 19;i ++){
                     printcells.printData(i, 0, readcells.Cliente[index++]);
                }
                
                indice_quantità = readcells.infoQuantità();
                indice_descrizione = readcells.infoDescrizione();
                indice_prezzounit = readcells.infoPrezzoUnit();
                IVA_value = readcells.infoIVA();
                
                for(int i = 0;i < indice_quantità;i ++){
                    printcells.printData(readcells.pos_quantità[i], 0, readcells.Quantità[i]);
                }
                for(int i = 0;i < indice_descrizione;i ++){
                    printcells.printData(readcells.pos_descrizione[i], 1, readcells.Descrizione[i]);
                }
                for(int i = 0;i < indice_prezzounit;i ++){
                    printcells.printData(readcells.pos_prezzounit[i], 4, readcells.PrezzoUnit[i]);
                }
                printcells.printData(38, 5, IVA_value);
                readcells.stampaDati();
                readcells.closeWorkbook();
    }
}
