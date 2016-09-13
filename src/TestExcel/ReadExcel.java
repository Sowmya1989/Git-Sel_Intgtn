package TestExcel;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel 
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
        String data = sh.getRow(0).getCell(0).getStringCellValue();
        System.out.println(data);
        String data1 = sh.getRow(0).getCell(1).getStringCellValue();
        System.out.println(data1);
        String data2 = sh.getRow(1).getCell(0).getStringCellValue();
        System.out.println(data2);
        String data3 = sh.getRow(1).getCell(1).getStringCellValue();
        System.out.println(data3);
        XSSFSheet sh1 = wb.getSheetAt(1);
        String SheetName1 = sh1.getSheetName();
        System.out.print("Sheet Name is "+SheetName1);
        System.out.println(" Data from sheet2 is given below");
        double data4 = sh1.getRow(0).getCell(0).getNumericCellValue();
        System.out.println(data4);
        double data5 = sh1.getRow(0).getCell(1).getNumericCellValue();
        System.out.println(data5);
        wb.close(); 
	}

}
