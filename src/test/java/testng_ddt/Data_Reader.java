package testng_ddt;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class Data_Reader 
{

	@Test //1st Test method
	(dataProvider="ReadVariant") //It get values from ReadVariant function method
	 
	//Here my all parameters from excel sheet are mentioned.
	public void AddVariants(String username, String password) throws Exception
	{
	//Data will set in Excel sheet once one parameter will get from excel sheet to Respective locator position.
	//DataSet++;
	System.out.println("user name is 		:" +username);
	System.out.println("password is			:" +password);
	}
	
	
	
	
	@DataProvider
	 public static Object[][] ReadVariant() throws IOException
	 {
	FileInputStream fis = new FileInputStream("C:/Users/bhanu/eclipse-workspace/testng_ddt/Data/ddt_Test.xlsx");
	 @SuppressWarnings("resource")
	XSSFWorkbook workbook = new XSSFWorkbook (fis); //get my workbook 
	 XSSFSheet sheet= workbook.getSheet("Sheet1");// get my sheet from workbook
	     //  XSSFRow Row=sheet.getRow(0);   //get my Row which start from 0   
			DataFormatter formatter = new DataFormatter();

	     int RowNum = sheet.getLastRowNum();// count my number of Rows
	    // int ColNum= Row.getLastCellNum(); // get last ColNum 
	     int ColNum = sheet.getRow(1).getLastCellNum();// get last ColNum
	     Object Data[][]= new Object[RowNum][ColNum]; // pass my  count data in array
	     
	     for(int i=0; i<RowNum; i++) //Loop work for Rows
	     {  
	    	 XSSFRow row= sheet.getRow(i+1);
	     
	    	 	for (int j=0; j<ColNum; j++) //Loop work for colNum
				    {
				     if(row==null)
				     Data[i][j]= "";
				     else 
				     {
					     XSSFCell cell= row.getCell(j);
					     if(cell==null)
					     Data[i][j]= ""; //if it get Null value it pass no data 
					     else
					     {
					     String value=formatter.formatCellValue(cell);
					     Data[i][j]=value; //This formatter get my all values as string i.e integer, float all type data value
				     }
				     }
				     }
	     }
	 
	     return Data;
	    }
		
}
