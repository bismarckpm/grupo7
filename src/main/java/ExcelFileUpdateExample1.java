import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import java.util.Scanner;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.DateUtil;

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Append new rows to an existing sheet.
 * 
 * @author www.codejava.net
 *
 */
public class ExcelFileUpdateExample1 {


	public static void main(String[] args) {
		String excelFilePath = "Inventario.xlsx";
                String ID;
                String Author;
                int Price;
                Scanner in = new Scanner(System.in);
                int Fila=-1;
                int Opcion=-1;
                
                System.out.println("Por favor ingrese el ID del campo a modificar");
                ID = in.nextLine();
                ID=ID+".0";
                
		
		try {
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
			Workbook workbook = WorkbookFactory.create(inputStream);

			Sheet sheet = workbook.getSheet("Java Books");
                        
                        int totalNoOfRows = sheet.getLastRowNum();
                        int totalNoOfColumns = sheet.getRow(0).getLastCellNum();
                                          
                        for (int rowIndex = 0; rowIndex <= totalNoOfRows; rowIndex++) {
                           
                            Row row = sheet.getRow(rowIndex);
                            if (row != null) {
                              Cell cell = row.getCell(0);
                              if (cell != null) {
                                  
                                String cellValue = null;

                                switch (cell.getCellType()) {
                                case Cell.CELL_TYPE_STRING:
                                        cellValue = cell.getStringCellValue();
                                        break;

                                case Cell.CELL_TYPE_FORMULA:
                                        cellValue = cell.getCellFormula();
                                        break;

                                case Cell.CELL_TYPE_NUMERIC:
                                        if (DateUtil.isCellDateFormatted(cell)) {
                                                cellValue = cell.getDateCellValue().toString();
                                        } else {
                                                cellValue = Double.toString(cell.getNumericCellValue());
                                        }
                                        break;

                                case Cell.CELL_TYPE_BLANK:
                                        cellValue = "";
                                        break;

                                case Cell.CELL_TYPE_BOOLEAN:
                                        cellValue = Boolean.toString(cell.getBooleanCellValue());
                                        break;

                                }
                                if (cellValue.equals(ID)){
                                    
                                    Fila = rowIndex;
                                    break;
                                }
                              }
                            }
                          }
                        
                         if (Fila==-1){
                             
                             System.out.println("Error, ID no encotrado.");
                             
                         }else{
                             
                             try{
                             
                             while((Opcion!=1)&&(Opcion!=2)){
                                 
                              System.out.println("Por favor marque 1 para cambiar el nombre del autor o 2 para cambiar el precio del libro, no se vale otro número diferente para las opciones a realizar.");
                              Opcion = in.nextInt();
                              
                             }
                             
                             if(Opcion==1){
                                 
                                 System.out.println("Por favor indique el nuevo nombre de Autor");
                                 Scanner author = new Scanner(System.in);
                                 Author = author.nextLine();
                                 Cell cell2Update = sheet.getRow(Fila).getCell(2);
                                 cell2Update.setCellValue(Author);
                                 
                             }
                             if(Opcion==2){
                                 
                                 System.out.println("Por favor indique el nuevo precio");
                                 Scanner price = new Scanner(System.in);
                                 Price = price.nextInt();
                                 Cell cell2Update = sheet.getRow(Fila).getCell(3);
                                 cell2Update.setCellValue(Price);
                                 
                             }
                             
                             System.out.println("Operanción exitosa, aquí podrá ver los datos contenidos en el archivo:");
                             
                             for (int rowIndex = 0; rowIndex <= totalNoOfRows; rowIndex++) {
                           
                                Row row = sheet.getRow(rowIndex);
                                
                                if (row != null) {
                                    
                                 for (int colIndex = 0; colIndex <= totalNoOfColumns; colIndex++) {
                                     
                                  Cell cell = row.getCell(colIndex);
                                  
                                  if (cell != null) {

                                    String cellValue = null;

                                    switch (cell.getCellType()) {
                                        
                                    case Cell.CELL_TYPE_STRING:
                                            cellValue = cell.getStringCellValue();
                                            break;

                                    case Cell.CELL_TYPE_FORMULA:
                                            cellValue = cell.getCellFormula();
                                            break;

                                    case Cell.CELL_TYPE_NUMERIC:
                                            if (DateUtil.isCellDateFormatted(cell)) {
                                                    cellValue = cell.getDateCellValue().toString();
                                            } else {
                                                    cellValue = Double.toString(cell.getNumericCellValue());
                                            }
                                            break;

                                    case Cell.CELL_TYPE_BLANK:
                                            cellValue = "";
                                            break;

                                    case Cell.CELL_TYPE_BOOLEAN:
                                            cellValue = Boolean.toString(cell.getBooleanCellValue());
                                            break;

                                    }

                                   
                                    System.out.println("["+cellValue+"]");


                                  }
                                }
                               }
                             }
                             
                           }catch(Exception e){
                               System.out.println("Ingresó un caracter diferente a los numeros 1 y 2, por favor corra de nuevo el programa y asegúrese de seguir los pasos correctamente");
                           }
                             
                         }
                        
			inputStream.close();

			FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
			
		} catch (IOException | EncryptedDocumentException
				| InvalidFormatException ex) {
			ex.printStackTrace();
		}
                
                
	}

}
