package br.com.xlsreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Reader {

	
	public void exportSheet(File file) throws IOException, EncryptedDocumentException, InvalidFormatException {
	
	    InputStream inp = new FileInputStream(file);
	    Workbook wb = WorkbookFactory.create(inp);
	    
	    Sheet sheet = wb.getSheetAt(0);
	    DataFormatter formatter = new DataFormatter();
	    
	    Path path = Paths.get(System.getProperty("user.dir"));
	    
	    String folderFileName = file.getName().split("\\.")[0];
	    File directory = new File(Paths.get(path + File.separator + "arquivos" + File.separator + folderFileName + File.separator).toString());
	    if(!directory.exists()) directory.mkdirs();
	    
	    Path p;
	    
	    for (Row row : sheet) {
	    	if(row.getRowNum() > 0) {
		    	if(formatter.formatCellValue(row.getCell(6)).equals("3") || formatter.formatCellValue(row.getCell(6)).equals("6")) {
		    		p = Paths.get(directory.getPath() + File.separator + formatter.formatCellValue(row.getCell(1)) + "_" + formatter.formatCellValue(row.getCell(4)) + "_" + formatter.formatCellValue(row.getCell(6)) + ".csv");
		    		StringBuilder sb = new StringBuilder().append(row.getRowNum()).append(";").append(formatter.formatCellValue(row.getCell(7))).append("\n");
		    		
		    		Files.write(p, sb.toString().getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
		    	}
	    		
//	    		System.out.println(folderFileName);
//	    		System.out.println(row.getRowNum() + " - " + formatter.formatCellValue(row.getCell(1)) + " - " + formatter.formatCellValue(row.getCell(4)) + " - " + formatter.formatCellValue(row.getCell(6)) + " - " + formatter.formatCellValue(row.getCell(7)));
	    	}


//	    	if(formatter.formatCellValue(row.getCell(6)).equals("3")) {
//	    		if(formatter.formatCellValue(row.getCell(4)).toUpperCase().equals("SUCESSO")) {
//	    			
//	    		} else if(formatter.formatCellValue(row.getCell(4)).toUpperCase().equals("INSUCESSO")) {
//	    			
//	    		}
//	    	} else if(formatter.formatCellValue(row.getCell(6)).equals("6")) {
//	    		if(formatter.formatCellValue(row.getCell(4)).toUpperCase().equals("SUCESSO")) {
//	    			
//	    		} else if(formatter.formatCellValue(row.getCell(4)).toUpperCase().equals("INSUCESSO")) {
//	    			
//	    		}
//	    	}
		}
	    inp.close();
	    wb.close();
	    
	    FileOutputStream fileOut = new FileOutputStream("workbook.xls");
	    wb.write(fileOut);
	    fileOut.close();
	}
	
}
