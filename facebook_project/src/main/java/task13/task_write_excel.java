package task13;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class task_write_excel {

	public static void main(String[] args) {
		
	
		task_write_excel obj = new task_write_excel();
		obj.writeExcelData();
		
		
	}
	
	
	public  void writeExcelData()
	{
	
		
		
			try {
				
				
				String outputPath = System.getProperty("user.dir") + "/task13.xlsx";
				System.out.println(outputPath);
				
				File file = new File(outputPath);
				
				FileOutputStream outstream = new FileOutputStream(outputPath);//open file
				
				
				XSSFWorkbook book = new XSSFWorkbook();//convert and open excel format
				
				XSSFSheet sheet = book.createSheet("newsheet");
				
				sheet.createRow(0).createCell(0).setCellValue("Name");
				sheet.createRow(1).createCell(0).setCellValue("John Doe");
				sheet.createRow(2).createCell(0).setCellValue("Jane Doe");
				sheet.createRow(3).createCell(0).setCellValue("Bob Smith");
				sheet.createRow(4).createCell(0).setCellValue("Swapnil");
				book.write(outstream);
				
				outstream.close();
				book.close();
				
				
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				
			}
			
			catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
	
			}
	}
}
	
	
	
	
	
	
