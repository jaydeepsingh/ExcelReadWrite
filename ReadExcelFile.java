import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadExcelFile {

	
	public void readExcel(String filePath, String fileName, String sheetName) throws IOException{
		
		File file = new File(filePath+"/"+fileName);
		
		//Scanner in = new Scanner(file);
		
		FileInputStream inputStream = new FileInputStream(file);
		
		Workbook testWorkbook = null;
		
		String fileExtensionName = fileName.substring(fileName.indexOf("."));
		
		System.out.println(fileExtensionName);
		
		if(fileExtensionName.equals(".xlsx")){
			testWorkbook = new XSSFWorkbook(inputStream);
			//testWorkbook = new XSSFWorkbook();
		}
		
		else if(fileExtensionName.equals(".xls")){
			 
	        //If it is xls file then create object of XSSFWorkbook class
	 
	       testWorkbook = new HSSFWorkbook(inputStream);
			//testWorkbook = new HSSFWorkbook();
	 
	    }
		System.out.println(testWorkbook);
		Sheet testSheet= testWorkbook.getSheet(sheetName);
		
		int rowCount = testSheet.getLastRowNum() - testSheet.getFirstRowNum();
		
		for (int i=0 ; i< rowCount+1 ; i++){
			Row row = testSheet.getRow(i);
			
			for(int j=0 ; j< row.getLastCellNum() ; j++){
				if(row.getCell(j).getCellType()==1){
				System.out.print(row.getCell(j).getStringCellValue() +"||");
				}else{
				System.out.print(row.getCell(j).getNumericCellValue());
				}
			}
			System.out.println();
		}
	}
	
	public static void main(String[] args) throws IOException{
		
		ReadExcelFile objExcelFile = new ReadExcelFile();
		
		String filePath = System.getProperty("user.dir")+"/src";
		//String filePath = System.getProperty("C:/Users/bhatijay/workspace/Test- ExcelReadWrite/src");
		objExcelFile.readExcel(filePath, "test.xlsx", "Sheet1");
		

	}

}
