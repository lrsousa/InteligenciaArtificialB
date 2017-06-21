package br.com.xlsreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Reader {

	
	public void exportSheet(File file) throws IOException, EncryptedDocumentException, InvalidFormatException {
	
	    InputStream inp = new FileInputStream(file);

	    Workbook wb = WorkbookFactory.create(inp);
	    Sheet sheet = wb.getSheetAt(0);
	    
	    for (Row row : sheet) {
			System.out.println(row.getCell(0));
			System.out.println(row.getCell(1));
		}
	    
	    Row row = sheet.getRow(2);
	    Cell cell = row.getCell(3);
	    if (cell == null)
	        cell = row.createCell(3);
	    cell.setCellType(CellType.STRING);
	    cell.setCellValue("a test");

	    // Write the output to a file
	    FileOutputStream fileOut = new FileOutputStream("workbook.xls");
	    wb.write(fileOut);
	    fileOut.close();
		
	}
}
