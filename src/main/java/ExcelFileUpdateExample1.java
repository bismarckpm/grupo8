import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
import java.util.Iterator;
import java.io.FileNotFoundException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Append new rows to an existing sheet.
 * 
 * @author www.codejava.net
 *
 */
public class ExcelFileUpdateExample1 {
	static Scanner scanner = new Scanner(System.in); 
	static int opcion = -1; 
	static String excelFilePath = "Inventario.xlsx";

	public static void main(String[] args) {
		menu(opcion);
	}

	public static void agregarRegistros(String excelFilePath){
		
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

	public static void mostrarTodosLosRegistros(String excelFilePath){
		try
		{
			FileInputStream file = new FileInputStream(new File(excelFilePath));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			//Iterate through sheets
			Iterator<Sheet> sheetIterator = workbook.iterator();
			while (sheetIterator.hasNext()) {
				XSSFSheet sheet = (XSSFSheet) sheetIterator.next();
					   
				//Iterate through each rows one by one
				Iterator<Row> rowIterator = sheet.iterator();
				while (rowIterator.hasNext()) 
				{
					Row row = rowIterator.next();
					//For each row, iterate through all the columns
					Iterator<Cell> cellIterator = row.cellIterator();
					
					while (cellIterator.hasNext()) 
					{
						Cell cell = cellIterator.next();
						//Check the cell type and format accordingly
						switch (cell.getCellType()) 
						{
							case Cell.CELL_TYPE_NUMERIC:
								int value = new Double(cell.getNumericCellValue()).intValue();
								System.out.print(value + " ");
								break;
							case Cell.CELL_TYPE_STRING:
								System.out.print(cell.getStringCellValue() + " ");
								break;
						}
					}
					System.out.println("");
				}
			}
			file.close();
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}

	}

public static void validarArchivo(String excelFilePath){
	File file = new File(excelFilePath);

	if (file.exists()){
		System.out.println("El archivo no existe. Creando uno");
		//hace lo del main
	}

	else{
		//crea el archivo y corre el main
	}

}
	public static void menu (int opcion){
		while(opcion != 0){
			try{
				System.out.println("Opciones:\n" +	
						"1.- Validar archivo existente\n" +
						"2.- Cantidad de registros por hoja\n" +
						"3.- Actualizacion de registro especifico\n" +
						"0.- Salir");
				opcion = Integer.parseInt(scanner.nextLine()); 
	
				switch(opcion){
				case 1: 
					System.out.println("1");
					validarArchivo(excelFilePath);
					agregarRegistros(excelFilePath);
					mostrarTodosLosRegistros(excelFilePath);
					break;
				case 2: 
					System.out.println("2");
					break;
				case 3: 
					System.out.println("3");
					break;
				case 0: 
					System.out.println("Adios!");
					break;
				default:
					System.out.println("Opcion no reconocido");break;
				}
				
				System.out.println("\n");
				
			}catch(Exception e){
				System.out.println("Uoop! Error!");
			}
		}

	}
}

