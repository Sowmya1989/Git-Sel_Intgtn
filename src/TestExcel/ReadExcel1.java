package TestExcel;

import java.io.File;
import java.io.ObjectInputStream.GetField;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel1 
{

	public static void main(String[] args) throws Exception
	{
		File src = new File("D:\\My World\\HMH\\Web Presence\\Automation\\CI_projects\\TestData.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(src);
		System.out.println("Active sheet index is "+wb.getActiveSheetIndex());
        XSSFSheet sh = wb.getSheetAt(0);
        String SheetName = sh.getSheetName();
        System.out.print("Sheet Name is "+SheetName);
        System.out.println(" Data from sheet1 is given below");
        
        wb.close(); 

	}

}
