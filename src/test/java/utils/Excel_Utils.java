package utils;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Excel_Utils {
	
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	
	public Excel_Utils(String excelPath, String sheetName) throws IOException
	{
		try {
		workbook = new XSSFWorkbook(excelPath);
		sheet = workbook.getSheet(sheetName);
		}
		catch(Exception exp)
		{
			System.out.println(exp.getCause());
			System.out.println(exp.getMessage());
			exp.printStackTrace();
		}
	}
	// funkce pro vraceni hodnoty ze zadané políčka
	public static int getCellData(int rowNum,int colNum) throws IOException {
		
		DataFormatter formatter = new DataFormatter();
		String value = formatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));
		
		int hod =-1;
		try {
			hod = Integer.parseInt(value);
		} catch (Exception e) {
			
		}
		return hod;
	}
	
	// funkce na řádky z excelu
	public static int getRowCount()
	{
	
		int rowCount = sheet.getPhysicalNumberOfRows();
		return rowCount;
	}
	
	// funkce na primarní čísla
	public static void getPrimeNumber(int val)
	{
		
		 int i,m=0,flag=0;      
		  m=val/2;
		  if(val>0) {
		  if(val==0||val==1){      
		  }else{  
		   for(i=2;i<=m;i++){      
		    if(val%i==0){         
		     flag=1;      
		     break;      
		    }      
		   }      
		   if(flag==0)  { System.out.println(val); }  
		  } 
		}
	}
	}
	

