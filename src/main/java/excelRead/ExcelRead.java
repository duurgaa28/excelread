package excelRead;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
	XSSFSheet sheet;
ExcelRead () throws IOException{
	FileInputStream inputfile =new FileInputStream("\\C:\\Users\\lenovo\\Downloads\\Book1.xlsx");
	XSSFWorkbook workbook=new XSSFWorkbook(inputfile);
	sheet=workbook.getSheet("Sheet1");
} 
public String readExcelData(int i,int j) {
	XSSFRow row=sheet.getRow(i);
	Cell cell=row.getCell(j);
	CellType type=cell.getCellType();
	
		switch(type){
		case NUMERIC:
			return String.valueOf(cell.getNumericCellValue());
		case STRING:
		return	cell.getStringCellValue();
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		}
		
	
		
	return "  ";
	}
	
public static void main (String args[]) throws IOException {
	ExcelRead ex=new ExcelRead();
	for(int i=0;i<6;i++) {
		for(int j=0;j<3;j++) {
			System.out.print(ex.readExcelData(i,j)+" ");
		}
		System.out.println();
	}

}
}