package Ress.Ress;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Exxx 
{
	public static void main(String[] args) throws IOException 
	{
		File loc = new File("E:\\64 bit\\Praveen_selenium\\first\\src\\main\\java\\first\\ExportExcel.xlsx");
		FileInputStream stream = new FileInputStream(loc);
		Workbook h = new XSSFWorkbook(stream);
		Workbook w = new XSSFWorkbook(stream);
		Sheet s=  w.getSheet("ExcelGuru99Demo");
		Row r =s.getRow(2);
		Cell c=r.getCell(1);
		System.out.println(c);
//		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) 
//		{
//			Row rr =s.getRow(i);
//			for (int j = 0; j < rr.getPhysicalNumberOfCells(); j++) 
//			{
//				Cell ca =rr.getCell(j);
//				int type =ca.getCellType();
//				if (type==1) 
//				{
//					String name =ca.getStringCellValue();
//					System.out.println(name);
//			
//				}
//				if (type==0) 
//				{ 
//					double d = ca.getNumericCellValue();
//					long l = (long) d;
//					String name = String.valueOf(l);
//					System.out.println(name);
//				}
//												
//			}
//			
//		}
		
	}

}
