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
import java.util.Arrays;
import java.util.HashMap;
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
	
	public void geraContagem(File file) throws IOException {
		List<String> spamFile = Files.readAllLines(file.toPath());
		System.out.println(file.getAbsolutePath());
		Path path = Paths.get(System.getProperty("user.dir"));
		
//		String pattern = Pattern.quote(System.getProperty("file.separator"));
		File directory = new File(Paths.get(path + File.separator + "finalSequencia" + File.separator + file.getParentFile().getName() + File.separator).toString());
		
		if(!directory.exists()) directory.mkdirs();

		
		Path p = Paths.get(directory.getPath() + File.separator + file.getName().replaceAll("FORA_", "FINAL_"));
		List<String> columnNameSequenciaDeEventos = new ArrayList<String>();
		
		HashMap<String, String> identificadorColunas = new HashMap<String, String>();
		HashMap<String, String> sequenciaComAlunos = new HashMap<String, String>();
		
		Path processadoComLinhasAlunos = Paths.get(path + File.separator + "processadosComLinhasAlunos" + File.separator +  file.getParentFile().getName() + File.separator + file.getName());
		List<String> arquivoComLinhasAlunos = Files.readAllLines(processadoComLinhasAlunos);
		HashMap<String, String> identificadorAlunos = new HashMap<String, String>();
		Integer i = 0;
		for(String s : arquivoComLinhasAlunos) {
			identificadorAlunos.put(i.toString(), s.split(";")[1]);
			i++;
		}
		HashSet<String> alunos = new HashSet<String>(); 

		i = 0;
		for(String s : spamFile) {
			String[] sSeparada = s.split(Pattern.quote("#SUP: "));
			String seq = sSeparada[0].replaceAll(" -1 ", ">");
			columnNameSequenciaDeEventos.add(seq);
			identificadorColunas.put(seq, i.toString());
			i++;
			
			String linhaAlunos = s.split(Pattern.quote("#SID: "))[1];

			sequenciaComAlunos.put(seq, linhaAlunos);
			
			List<String> alunosV = Arrays.asList(linhaAlunos.split(" "));
			for (String a : alunosV) {
				alunos.add(identificadorAlunos.get(a));
			}
		}
		
		StringBuilder sb = new StringBuilder().append("aluno;");
		for(Iterator<String> iterator = columnNameSequenciaDeEventos.iterator(); iterator.hasNext();) {
			sb.append(iterator.next()).append(";");
		}
		sb.append("\n");
		Files.write(p, sb.toString().getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
		
		
		for (String aluno : alunos) {
			sb = new StringBuilder().append(aluno).append(";");
			for (String col : columnNameSequenciaDeEventos) {
				List<String> l = Arrays.asList(sequenciaComAlunos.get(col).split(" "));
				Boolean consta = false;
				int valor = 0;
				for (String a : l) {
					if(identificadorAlunos.get(a).equals(aluno)) {
						consta = true;
						valor++;
//						System.out.println("consta");
					}
//					System.out.println(a);
				}
				
				if(consta) {
//					System.out.println(aluno);
					sb.append(valor).append(";");
				} else {
					sb.append(";");
				}
				
				
			}
			sb.append("\n");
//			System.out.println(sb.toString());
			Files.write(p, sb.toString().getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
			
		}
		
	}
	
	public void geraContagemIndividual(File file) throws IOException {
		List<String> spamFile = Files.readAllLines(file.toPath());
		System.out.println(file.getAbsolutePath());
		Path path = Paths.get(System.getProperty("user.dir"));
		
//		String pattern = Pattern.quote(System.getProperty("file.separator"));
		File directory = new File(Paths.get(path + File.separator + "finalSequencia" + File.separator + "spam" + File.separator + file.getParentFile().getName() + File.separator).toString());
		
		if(!directory.exists()) directory.mkdirs();
		
		Path p = Paths.get(directory.getPath() + File.separator + file.getName());
		List<String> columnNameSequenciaDeEventos = new ArrayList<String>();
		
		HashMap<String, String> identificadorColunas = new HashMap<String, String>();
		HashMap<String, String> sequenciaComAlunos = new HashMap<String, String>();
		
		Path processadoComLinhasAlunos = Paths.get(path + File.separator + "processadosComLinhasAlunos" + File.separator +  file.getParentFile().getName() + File.separator + "FORA_" + file.getName());
		List<String> arquivoComLinhasAlunos = Files.readAllLines(processadoComLinhasAlunos);
		HashMap<String, String> identificadorAlunos = new HashMap<String, String>();
		Integer i = 0;
		for(String s : arquivoComLinhasAlunos) {
			identificadorAlunos.put(i.toString(), s.split(";")[1]);
			i++;
		}
		HashSet<String> alunos = new HashSet<String>(); 

		i = 0;
		for(String s : spamFile) {
			String[] sSeparada = s.split(Pattern.quote("#SUP: "));
			String seq = sSeparada[0].replaceAll(" -1 ", ">");
			columnNameSequenciaDeEventos.add(seq);
			identificadorColunas.put(seq, i.toString());
			i++;
			
			String linhaAlunos = s.split(Pattern.quote("#SID: "))[1];

			sequenciaComAlunos.put(seq, linhaAlunos);
			
			List<String> alunosV = Arrays.asList(linhaAlunos.split(" "));
			for (String a : alunosV) {
				alunos.add(identificadorAlunos.get(a));
			}
		}
		
		StringBuilder sb = new StringBuilder().append("aluno;");
		for(Iterator<String> iterator = columnNameSequenciaDeEventos.iterator(); iterator.hasNext();) {
			sb.append(iterator.next()).append(";");
		}
		sb.append("\n");
		Files.write(p, sb.toString().getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
		
		
		for (String aluno : alunos) {
			sb = new StringBuilder().append(aluno).append(";");
			for (String col : columnNameSequenciaDeEventos) {
				List<String> l = Arrays.asList(sequenciaComAlunos.get(col).split(" "));
				Boolean consta = false;
				int valor = 0;
				for (String a : l) {
					if(identificadorAlunos.get(a).equals(aluno)) {
						consta = true;
						valor++;
//						System.out.println("consta");
					}
//					System.out.println(a);
				}
				
				if(consta) {
//					System.out.println(aluno);
					sb.append(valor).append(";");
				} else {
					sb.append(";");
				}
				
				
			}
			sb.append("\n");
//			System.out.println(sb.toString());
			Files.write(p, sb.toString().getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
			
		}
		
	}

	
	public void geraContagemV2(File file) throws IOException {
		List<String> spamFile = Files.readAllLines(file.toPath());
		System.out.println(file.getAbsolutePath());
		Path path = Paths.get(System.getProperty("user.dir"));
		
		String[] vetorNomeArquivo = file.getName().split("_");
		StringBuilder novoNomeArquivo = new StringBuilder().append(vetorNomeArquivo[0]).append("_").append(vetorNomeArquivo[1]).append("_").append("SUC_INSUC").append("_").append(vetorNomeArquivo[3]);
		
//		String pattern = Pattern.quote(System.getProperty("file.separator"));
		File directory = new File(Paths.get(path + File.separator + "finalSequenciaV2" + File.separator + file.getParentFile().getName() + File.separator).toString());
		
		if(!directory.exists()) directory.mkdirs();

		Path p = Paths.get(directory.getPath() + File.separator + novoNomeArquivo.toString());
		
		Path processadoComLinhasAlunos = Paths.get(path + File.separator + "processadosComLinhasAlunos" + File.separator +  file.getParentFile().getName() + File.separator + file.getName());
		List<String> arquivoComLinhasAlunos = Files.readAllLines(processadoComLinhasAlunos);
		HashMap<String, String> identificadorAlunos = new HashMap<String, String>();
		Integer i = 0;
		for(String s : arquivoComLinhasAlunos) {
			identificadorAlunos.put(i.toString(), s.split(";")[1]);
			i++;
		}
		i = 0;
		StringBuilder sb = new StringBuilder().append("aluno;");
		for(String s : spamFile) {
			String[] sSeparada = s.split(Pattern.quote("#SUP: "));
			String seq = sSeparada[0].replaceAll(" -1 ", ">");
			String linhaAlunos = s.split(Pattern.quote("#SID: "))[1];
			
			List<String> alunosV = Arrays.asList(linhaAlunos.split(" "));
			for (String a : alunosV) {
				Files.write(p, (identificadorAlunos.get(a) + ";" + seq + "\n").getBytes(), StandardOpenOption.APPEND, StandardOpenOption.CREATE);
			}
		}
		
	}

	
}
