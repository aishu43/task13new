package task13;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class task_read_excel {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		task_read_excel obj = new 	task_read_excel();
		obj.readExcelData();
		
		
	}
	
	
	public  void readExcelData()
	{
	
		
		
			try {
				
				
				String inputPath = System.getProperty("user.dir") + "/task13.xlsx";
				System.out.println(inputPath);
				
				FileInputStream instream = new FileInputStream(inputPath);//open file
				
				
				XSSFWorkbook book = new XSSFWorkbook(instream);//convert and open excel format
				XSSFSheet sheet = book.getSheet("Sheet1");
				
				
				Row row = sheet.getRow(0);
				Cell c = row.getCell(0);
				System.out.println(sheet.getRow(0).getCell(0));
				Row row1 = sheet.getRow(1);
				Cell c1 = row1.getCell(1);
				System.out.println(sheet.getRow(0).getCell(1));
				Row row2 = sheet.getRow(2);
				Cell c2 = row2.getCell(2);
				System.out.println(sheet.getRow(1).getCell(2));
				Row row3 = sheet.getRow(3);
				Cell c3 = row3.getCell(3);
				System.out.println(sheet.getRow(2).getCell(3));
				
				/*for (int i=0;i<1;i++)
				{
					Cell Cell2 = sheet.getRow(2).getCell(0);
					System.out.println(Cell2);*/
					
					book.close();
					instream.close();
				}
					
					/*int lastRow = sheet.getLastRowNum();
					
					for (int j=1; j<=lastRow;j++)
					{
						Row row11=sheet.getRow(i);
						Cell cellUser = row11.getCell(1);
						Cell cellPass = row11.getCell(2);
						
						System.out.println(cellUser +":"+ cellPass);
					}}}*/
						 ///entire excel row and col will print
					
				
				
			 catch (FileNotFoundException e) 
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
				
			}
			
			
			catch (IOException e)
			{
					// TODO Auto-generated catch block
					e.printStackTrace();
					
			}
			
			
			}
}