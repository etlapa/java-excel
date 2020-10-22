package mx.edev.java.excel.writer;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import mx.edev.java.excel.ExcelExecution;

public class ExcelWriter2003 implements ExcelExecution {

	public void exec(String outputFile) {
		try (FileOutputStream fileStream = new FileOutputStream(outputFile)) {
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet();
			// Create First Row
			HSSFRow row1 = sheet.createRow(0);
			HSSFCell r1c1 = row1.createCell(0);
			r1c1.setCellValue("Emd Id");
			HSSFCell r1c2 = row1.createCell(1);
			r1c2.setCellValue("NAME");
			HSSFCell r1c3 = row1.createCell(2);
			r1c3.setCellValue("AGE");
			// Create Second Row
			HSSFRow row2 = sheet.createRow(1);
			HSSFCell r2c1 = row2.createCell(0);
			r2c1.setCellValue("1");
			HSSFCell r2c2 = row2.createCell(1);
			r2c2.setCellValue("Ram");
			HSSFCell r2c3 = row2.createCell(2);
			r2c3.setCellValue("20");
			// Create Third Row
			HSSFRow row3 = sheet.createRow(2);
			HSSFCell r3c1 = row3.createCell(0);
			r3c1.setCellValue("2");
			HSSFCell r3c2 = row3.createCell(1);
			r3c2.setCellValue("Shyam");
			HSSFCell r3c3 = row3.createCell(2);
			r3c3.setCellValue("25");
			workbook.write(fileStream);
			System.out.println("Excel file was created at: " + outputFile);
		} catch (IOException e) {
			System.err.println("Could not create XLSX file at " + outputFile);
			e.printStackTrace();
		}
	}
}