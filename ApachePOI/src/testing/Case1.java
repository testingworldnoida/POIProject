package testing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Case1 {

	public static void main(String aa[]) 
	
	{
		try
		{
			File f = new File("G:\\Data3.xlsx");
			FileOutputStream fo= new FileOutputStream(f);
			XSSFWorkbook wk = new XSSFWorkbook();
			XSSFSheet s1 = wk.createSheet("ABC");
			Row r1 =s1.createRow(0);
			Cell c = r1.createCell(0);
			c.setCellValue("DfdsfSD");
			wk.write(fo);
			fo.flush();
			fo.close();
			System.out.println("HELLO");

		}catch(Exception e){System.out.println(e.getMessage());}
	}
	
}
