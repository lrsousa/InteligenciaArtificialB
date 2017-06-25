package br.com.xlsreader;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class Main {
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		Path path = Paths.get(System.getProperty("user.dir"));
		Path destinyFolder = Paths.get(path.toString() + File.separator + "backup" + File.separator);

		File folder = new File(path.toString() + File.separator + "entrada" + File.separator);
		
		Reader r = new Reader();

//		for (File f : folder.listFiles()) {
//			System.out.println("Lendo arquivo: " + f.getName());
//			r.separaPeriodos(f);
//			r.separaIndividual(f);
//			r.separaPeriodosComLinhasAlunos(f);
			
//			Files.move(f.toPath(), destinyFolder.resolve(f.getName()));
			
//		}
		folder = new File(path.toString() + File.separator + "processados" + File.separator + "spam" + File.separator);
		for (File f : folder.listFiles()) {
			for(File file : f.listFiles()) {
				r.juntaEssaPorraToda(file);
			}
		}
		
		System.out.println("Fim de leitura.");
	}
}
