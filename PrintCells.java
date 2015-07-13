/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package printcustom;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

/**
 *
 * @author ebissolo
 */
public class PrintCells {
    
    Cell cell1, cell2;
    
    public void printData(int row, int col, String data ) throws FileNotFoundException, IOException {
         FileInputStream fsIP= new FileInputStream(new File("Ricevuta_campi_vuoti.xls")); 
         HSSFWorkbook wb = new HSSFWorkbook(fsIP);               
         HSSFSheet worksheet = wb.getSheetAt(0);                          
         cell1 = worksheet.getRow(row).getCell(col);  
         cell2 = worksheet.getRow(row).getCell(col+8);
         cell1.setCellValue(data);
         cell2.setCellValue(data);
         fsIP.close(); 
         FileOutputStream output_file =new FileOutputStream(new File("Ricevuta_campi_vuoti.xls"));       
         wb.write(output_file);           
         output_file.close(); 
    }
}

