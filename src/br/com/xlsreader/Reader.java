package br.com.xlsreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.regex.Pattern;
import java.util.stream.Stream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Reader {

	
	public void separaPeriodos(File file) throws IOException, EncryptedDocumentException, InvalidFormatException {
	
	    InputStream inp = new FileInputStream(file);
	    Workbook wb = WorkbookFactory.create(inp);
	    
	    Sheet sheet = wb.getSheetAt(0);
	    DataFormatter formatter = new DataFormatter();
	    
	    Set<String> semestres = new HashSet<String>();
	    for (Row row : sheet) {
	    	if(row.getRowNum() > 0) {
	    		semestres.add(formatter.formatCellValue(row.getCell(1)));
	    	}
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
	    			if(!formatter.formatCellValue(row.getCell(1)).equals(s)) {
	    				if(formatter.formatCellValue(row.getCell(6)).equals("1") || formatter.formatCellValue(row.getCell(6)).equals("2") || formatter.formatCellValue(row.getCell(6)).equals("3")) {
	    					p = Paths.get(directory.getPath() + File.separator + "FORA_" + s + "_" + formatter.formatCellValue(row.getCell(4)).toUpperCase() + "_3.csv");
//			    			StringBuilder sb = new StringBuilder().append(row.getRowNum()).append(";").append(formatter.formatCellValue(row.getCell(7))).append("\n"); //com numeros de linha
	    					StringBuilder sb = new StringBuilder().append(formatter.formatCellValue(row.getCell(7))).append("\n");
	    					
	    					Files.write(p, sb.toString().getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
	    				}
	    				if(formatter.formatCellValue(row.getCell(6)).equals("1") || formatter.formatCellValue(row.getCell(6)).equals("2") || formatter.formatCellValue(row.getCell(6)).equals("3") || formatter.formatCellValue(row.getCell(6)).equals("4") || formatter.formatCellValue(row.getCell(6)).equals("5") || formatter.formatCellValue(row.getCell(6)).equals("6")) {
	    					p = Paths.get(directory.getPath() + File.separator + "FORA_" + s + "_" + formatter.formatCellValue(row.getCell(4)).toUpperCase() + "_6.csv");
//				    		StringBuilder sb = new StringBuilder().append(row.getRowNum()).append(";").append(formatter.formatCellValue(row.getCell(7))).append("\n"); //com numeros de linha
	    					StringBuilder sb = new StringBuilder().append(formatter.formatCellValue(row.getCell(7))).append("\n");
	    					
	    					Files.write(p, sb.toString().getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
	    				}
	    			}
	    		}
	    	}
	    }
	    inp.close();
	    wb.close();
	}
	
	public void separaIndividual(File file) throws IOException, EncryptedDocumentException, InvalidFormatException {
		
	    InputStream inp = new FileInputStream(file);
	    Workbook wb = WorkbookFactory.create(inp);
	    
	    Sheet sheet = wb.getSheetAt(0);
	    DataFormatter formatter = new DataFormatter();
	    
	    Set<String> semestres = new HashSet<String>();
	    for (Row row : sheet) {
	    	if(row.getRowNum() > 0) {
	    		semestres.add(formatter.formatCellValue(row.getCell(1)));
	    	}
	    }
	    
	    Path path = Paths.get(System.getProperty("user.dir"));
	    
	    String folderFileName = file.getName().split("\\.")[0];
	    File directory = new File(Paths.get(path + File.separator + "individual" + File.separator + folderFileName + File.separator).toString());
	    
	    System.out.println("Criando diretório: " + directory.getAbsolutePath());
	    if(!directory.exists()) directory.mkdirs();
	    
	    Path p;
	    for(String s : semestres) {
	    	for (Row row : sheet) {
	    		if(row.getRowNum() > 0) {
	    			if(formatter.formatCellValue(row.getCell(1)).equals(s)) {
	    				if(formatter.formatCellValue(row.getCell(6)).equals("1") || formatter.formatCellValue(row.getCell(6)).equals("2") || formatter.formatCellValue(row.getCell(6)).equals("3")) {
	    					p = Paths.get(directory.getPath() + File.separator + s + "_" + formatter.formatCellValue(row.getCell(4)).toUpperCase() + "_3.csv");
//			    			StringBuilder sb = new StringBuilder().append(row.getRowNum()).append(";").append(formatter.formatCellValue(row.getCell(7))).append("\n"); //com numeros de linha
	    					StringBuilder sb = new StringBuilder().append(formatter.formatCellValue(row.getCell(7))).append("\n");
	    					
	    					Files.write(p, sb.toString().getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
	    				}
	    				if(formatter.formatCellValue(row.getCell(6)).equals("1") || formatter.formatCellValue(row.getCell(6)).equals("2") || formatter.formatCellValue(row.getCell(6)).equals("3") || formatter.formatCellValue(row.getCell(6)).equals("4") || formatter.formatCellValue(row.getCell(6)).equals("5") || formatter.formatCellValue(row.getCell(6)).equals("6")) {
	    					p = Paths.get(directory.getPath() + File.separator + s + "_" + formatter.formatCellValue(row.getCell(4)).toUpperCase() + "_6.csv");
//				    		StringBuilder sb = new StringBuilder().append(row.getRowNum()).append(";").append(formatter.formatCellValue(row.getCell(7))).append("\n"); //com numeros de linha
	    					StringBuilder sb = new StringBuilder().append(formatter.formatCellValue(row.getCell(7))).append("\n");
	    					
	    					Files.write(p, sb.toString().getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
	    				}
	    			}
	    		}
	    	}
	    }
	    inp.close();
	    wb.close();
	}

	public void separaPeriodosComLinhasAlunos(File file) throws EncryptedDocumentException, InvalidFormatException, IOException {
		InputStream inp = new FileInputStream(file);
	    Workbook wb = WorkbookFactory.create(inp);
	    
	    Sheet sheet = wb.getSheetAt(0);
	    DataFormatter formatter = new DataFormatter();
	    
	    Set<String> semestres = new HashSet<String>();
	    for (Row row : sheet) {
	    	if(row.getRowNum() > 0) {
	    		semestres.add(formatter.formatCellValue(row.getCell(1)));
	    	}
	    }
	    
	    Path path = Paths.get(System.getProperty("user.dir"));
	    
	    String folderFileName = file.getName().split("\\.")[0];
	    File directory = new File(Paths.get(path + File.separator + "processadosComLinhasAlunos" + File.separator + folderFileName + File.separator).toString());
	    
	    System.out.println("Criando diretório: " + directory.getAbsolutePath());
	    if(!directory.exists()) directory.mkdirs();
	    
	    Path p;
	    for(String s : semestres) {
	    	for (Row row : sheet) {
	    		if(row.getRowNum() > 0) {
	    			if(!formatter.formatCellValue(row.getCell(1)).equals(s)) {
	    				if(formatter.formatCellValue(row.getCell(6)).equals("1") || formatter.formatCellValue(row.getCell(6)).equals("2") || formatter.formatCellValue(row.getCell(6)).equals("3")) {
	    					p = Paths.get(directory.getPath() + File.separator + "FORA_" + s + "_" + formatter.formatCellValue(row.getCell(4)).toUpperCase() + "_3.csv");
			    			StringBuilder sb = new StringBuilder().append(row.getRowNum()).append(";").append(formatter.formatCellValue(row.getCell(3))).append(";").append(formatter.formatCellValue(row.getCell(7))).append("\n");
	    					Files.write(p, sb.toString().getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
	    				}
	    				if(formatter.formatCellValue(row.getCell(6)).equals("1") || formatter.formatCellValue(row.getCell(6)).equals("2") || formatter.formatCellValue(row.getCell(6)).equals("3") || formatter.formatCellValue(row.getCell(6)).equals("4") || formatter.formatCellValue(row.getCell(6)).equals("5") || formatter.formatCellValue(row.getCell(6)).equals("6")) {
	    					p = Paths.get(directory.getPath() + File.separator + "FORA_" + s + "_" + formatter.formatCellValue(row.getCell(4)).toUpperCase() + "_6.csv");
	    					StringBuilder sb = new StringBuilder().append(row.getRowNum()).append(";").append(formatter.formatCellValue(row.getCell(3))).append(";").append(formatter.formatCellValue(row.getCell(7))).append("\n");
	    					Files.write(p, sb.toString().getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
	    				}
	    			}
	    		}
	    	}
	    }
	    inp.close();
	    wb.close();
		
	}
	
	public void juntaEssaPorraToda(File file) throws IOException {
		List<String> lines = Files.readAllLines(file.toPath());
		System.out.println(file.getAbsolutePath());
		Path path = Paths.get(System.getProperty("user.dir"));
		
		String pattern = Pattern.quote(System.getProperty("file.separator"));
		File directory = new File(Paths.get(path + File.separator + "finalSequencia" + File.separator + file.getParentFile().getName() + File.separator).toString());
		
		if(!directory.exists()) directory.mkdirs();

		System.out.println(directory);
		
		Path p = Paths.get(directory.getPath() + File.separator + file.getName().replaceAll("FORA_", "FINAL_"));
//		Files.write(p, "OPAAAAA".getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
		
		List<String> columnNameSequenciaDeEventos = new ArrayList<String>();
		lines.forEach(s -> {columnNameSequenciaDeEventos.add(s.split(Pattern.quote(" #SUP: "))[0].replaceAll(" -1 -2 ", ">").replaceAll(" -1", ">").replaceAll(" ", ""));});
		
		StringBuilder sb = new StringBuilder().append("aluno;");
		for(Iterator<String> iterator = columnNameSequenciaDeEventos.iterator(); iterator.hasNext();) {
			sb.append(iterator.next()).append(";");
			;
		}
		Files.write(p, sb.toString().getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
		
		System.out.println(sb.toString());
		for (String c : columnNameSequenciaDeEventos) {
			System.out.println(c);
		}
		
	}
}
