package br.com.xlsreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.HashSet;
import java.util.Set;

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
	    
	    Set<String> semestres = new HashSet<String>();
	    for (Row row : sheet) {
	    	semestres.add(formatter.formatCellValue(row.getCell(1)));
	    }
	    
	    Path path = Paths.get(System.getProperty("user.dir"));
	    
	    String folderFileName = file.getName().split("\\.")[0];
	    File directory = new File(Paths.get(path + File.separator + "processados" + File.separator + folderFileName + File.separator).toString());
	    
	    System.out.println("Criando diretório: " + directory.getAbsolutePath());
	    if(!directory.exists()) directory.mkdirs();
	    
	    Path p;
	    for(String s : semestres) {
	    	for (Row row : sheet) {
	    		if(row.getRowNum() > 0) {
			    	if((formatter.formatCellValue(row.getCell(6)).equals("3") || formatter.formatCellValue(row.getCell(6)).equals("6")) && !formatter.formatCellValue(row.getCell(1)).equals(s)) {
			    		p = Paths.get(directory.getPath() + File.separator + "FORA_" + s + "_" + formatter.formatCellValue(row.getCell(4)).toUpperCase() + "_" + formatter.formatCellValue(row.getCell(6)) + ".csv");
			    		StringBuilder sb = new StringBuilder().append(formatter.formatCellValue(row.getCell(7))).append("\n");
			    		
			    		Files.write(p, sb.toString().getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
			    	}
	    		}
	    	}
	    }
	    inp.close();
	    wb.close();
	}
	
}
