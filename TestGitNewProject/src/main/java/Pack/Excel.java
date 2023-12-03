package Pack;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.dev.XSSFSave;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {;
	XSSFSheet sh;
	
	public Excel() throws IOException {
		FileInputStream f=new FileInputStream("C:\\Users\\jobin\\OneDrive\\Documents\\Jessin John\\test.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(f);
	    sh=wb.getSheet("Sheet1");
	}

	public String readData(int a,int b) {
		Row r=sh.getRow(a);
		Cell c=r.getCell(b);
		int cellType=c.getCellType();//0 or 1
		switch (cellType) {
		case Cell.CELL_TYPE_NUMERIC:
			double aa=c.getNumericCellValue();
			return String.valueOf(c);
		case Cell.CELL_TYPE_STRING:
			return c.getStringCellValue();
		}
		return null;
	}
}
