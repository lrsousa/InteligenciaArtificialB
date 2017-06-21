package br.com.xlsreader;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class Main {
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		Path path = Paths.get(System.getProperty("user.dir"));
		File folder = new File(path.toString() + File.separator + "arquivos" + File.separator);
		
		Reader r = new Reader();

		for (File f : folder.listFiles()) {
			System.out.println("Lendo arquivo: " + f.getName());
			r.exportSheet(f);
		}
	}
}
