package br.com.xlsreader;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class Main {
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		Path path = Paths.get(System.getProperty("user.dir"));
		File f = new File(path.toString() + File.separator + "teste.xls");

		Reader r = new Reader();
		
		r.exportSheet(f);
		
		
	}
}
