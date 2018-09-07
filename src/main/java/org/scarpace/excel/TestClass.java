package org.scarpace.excel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;



public class TestClass {
	
	public static void main(String [] args) throws IOException, InvalidFormatException{
		
		String s = "C:\\StateOfVirginia\\Crystal Setup Files\\RP.144FastTrackReport.xls";
		
		File file = new File(s);

	FileInputStream fis = new FileInputStream("C:\\StateOfVirginia\\Crystal Setup Files\\RP.144FastTrackReport.xls");
	
	//Workbook wb = WorkbookFactory.create(fis);
	
	HSSFWorkbook wb = new HSSFWorkbook(fis);
	
	//Workbook wb = new HSSFWorkbook(fis);
	
	if (file.isFile() && file.exists()) {
		System.out.println("hurray! We've just opened a workbook");
	} else {
		System.out.println("Ahh! there was an error. Please make sure that the file path is correct.");
	}
	
	HSSFSheet sheet = wb.getSheetAt(0);
	//HSSFRow row = sheet.getRow(0);
	HSSFCell cell = sheet.getRow(0).getCell(0);//row.getCell(1);
	
	CellStyle cs = wb.createCellStyle();//wb.createCellStyle();
	cs.setFillForegroundColor(IndexedColors.BLUE.getIndex());
	cs.setFillPattern(CellStyle.SOLID_FOREGROUND); //FillPatternType.SOLID_FOREGROUND
	cell.setCellStyle(cs);
	
	FileOutputStream out = new FileOutputStream
			(new File("C:\\StateOfVirginia\\Crystal Setup Files\\RP.144FastTrackReport.xls"));
    wb.write(out);
    out.close();
	fis.close();
	
	/* ByteArrayOutputStream bout=new ByteArrayOutputStream();
		// writing the modified content from workbook to the ByteArrayOutputStream instance
	 wb.write(bout);
		bout.close();
	    fis.close();*/


	
	}
}