import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelFile {

	public void writeExcel(String filePath, String fileName, String sheetName,String[] dataToWrite)
			throws IOException {

		File file = new File(filePath + "/" + fileName);

		// Scanner in = new Scanner(file);

		FileInputStream inputStream = new FileInputStream(file);

		Workbook testWorkbook = null;

		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		System.out.println(fileExtensionName);

		if (fileExtensionName.equals(".xlsx")) {
			testWorkbook = new XSSFWorkbook(inputStream);
			// testWorkbook = new XSSFWorkbook();
		}

		else if (fileExtensionName.equals(".xls")) {

			// If it is xls file then create object of XSSFWorkbook class

			testWorkbook = new HSSFWorkbook(inputStream);
			// testWorkbook = new HSSFWorkbook();

		}
		System.out.println(testWorkbook);
		Sheet testSheet = testWorkbook.getSheet(sheetName);

		int rowCount = testSheet.getLastRowNum() - testSheet.getFirstRowNum();
		
		Row row = testSheet.getRow(0);
		Row newRow = testSheet.createRow(rowCount+1);
		
	    for(int j = 0; j < row.getLastCellNum(); j++){
	    	 
	        //Fill data in row
	 
	        Cell cell = newRow.createCell(j);
	 
	        cell.setCellValue(dataToWrite[j]);
	 
	    }
	 
	    //Close input stream
	 
	    inputStream.close();
	 
	    //Create an object of FileOutputStream class to create write data in excel file
	 
	    FileOutputStream outputStream = new FileOutputStream(file);
	 
	    //write data in the excel file
	 
	    testWorkbook.write(outputStream);
	 
	    //close output stream
	 
	    outputStream.close();
		
		
	}

	public static void main(String[] args) throws IOException{
		  //Create an array with the data in the same order in which you expect to be filled in excel file
		 
        String[] valueToWrite = {"d","4"};
 
        //Create an object of current class
 
        WriteExcelFile objExcelFile = new WriteExcelFile();
 
        //Write the file using file name , sheet name and the data to be filled
        String filePath = System.getProperty("user.dir")+"/src";
        objExcelFile.writeExcel(filePath,"test.xlsx","Sheet1",valueToWrite);
 
    }
	

}
