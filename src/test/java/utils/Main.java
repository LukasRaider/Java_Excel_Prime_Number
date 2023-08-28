package utils;

import java.io.IOException;

import org.apache.commons.compress.harmony.unpack200.bytecode.ClassFile;

import java.io.File;

public class Main {

	
	public static void main(String[] args) throws IOException{
		
		if (args.length==0) {
			System.out.println("Proved nacteni souboru z prikazoveho radku.");
			return;
		}
		
		//String excelPath = "./data/vzorek_dat.xlsx";
		String filePath = args[0];
		File file = new File(filePath);
		
		if(file.exists()) {
			System.out.println("Soubor existuje na dannem miste.");
			if(file.isFile()) {
				System.out.println("To je bezny soubor.");
			}else if(file.isDirectory()) {
				System.out.println("To je adresar.");
			}
			else {
				System.out.println("Zadany soubor neexistuje.");
			}
		}
		
		String sheetName = "List1";
		Excel_Utils excel_Utils = new Excel_Utils(filePath, sheetName);
		
		int Numrow=excel_Utils.getRowCount();
	
		for (int i = 1; i < Numrow; i++) {
			int value = excel_Utils.getCellData(i,1);
			
			if (value != -1)
			{
				excel_Utils.getPrimeNumber(value);
			}
			
	
		}
		
	}

	
}
