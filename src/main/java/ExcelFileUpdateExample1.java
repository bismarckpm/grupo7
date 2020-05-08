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

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Append new rows to an existing sheet.
 * 
 * @author www.codejava.net
 *
 */
public class ExcelFileUpdateExample1 {

	public static void existe(){
        String excelFilePath = "Inventario.xlsx";
        try{

            File file = new File(excelFilePath);
            if(!file.exists()){
                Workbook workbook = new XSSFWorkbook();
                Sheet sheet = workbook.createSheet("Inventario");

                FileOutputStream fileOut = new FileOutputStream(excelFilePath);
                workbook.write(fileOut);

                fileOut.close();


                FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
                workbook = WorkbookFactory.create(inputStream);

                sheet = workbook.getSheetAt(0);

                Row row = sheet.createRow(0);

                row.createCell(0).setCellValue("No");
                row.createCell(1).setCellValue("Book Title");
                row.createCell(2).setCellValue("Author");
                row.createCell(3).setCellValue("Price");

                inputStream.close();

                FileOutputStream outputStream = new FileOutputStream(excelFilePath);
                workbook.write(outputStream);
                workbook.close();
                outputStream.close();
            }

        }catch(FileNotFoundException ex){
            ex.printStackTrace();
        }
        catch(IOException ex){
            ex.printStackTrace();
        }
        catch(InvalidFormatException ex){
            ex.printStackTrace();
        }
    }


	public static void main(String[] args) {
		String excelFilePath = "Inventario.xlsx";
		existe();
		
		try {
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
			Workbook workbook = WorkbookFactory.create(inputStream);

			Sheet sheet = workbook.getSheetAt(0);

			Object[][] bookData = {
					{"El que se duerme pierde", "Tom Peter", 16},
					{"Sin lugar a duda", "Ana Gutierrez", 26},
					{"El arte de dormir", "Nico", 32},
					{"Buscando a Nemo", "Humble Po", 41},
			};

			int rowCount = sheet.getLastRowNum();

			for (Object[] aBook : bookData) {
				Row row = sheet.createRow(++rowCount);

				int columnCount = 0;
				
				Cell cell = row.createCell(columnCount);
				cell.setCellValue(rowCount);
				
				for (Object field : aBook) {
					cell = row.createCell(++columnCount);
					if (field instanceof String) {
						cell.setCellValue((String) field);
					} else if (field instanceof Integer) {
						cell.setCellValue((Integer) field);
					}
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
