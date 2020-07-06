package mx.edev.java.excel.writer;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import mx.edev.java.excel.ExcelExecution;

public class ExcelWriter implements ExcelExecution {

	public void exec(String outputFile) {
		try (FileOutputStream fileStream = new FileOutputStream(outputFile)) {
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet();
			// Create First Row
			XSSFRow row1 = sheet.createRow(0);
			XSSFCell r1c1 = row1.createCell(0);
			r1c1.setCellValue("Emd Id");
			XSSFCell r1c2 = row1.createCell(1);
			r1c2.setCellValue("NAME");
			XSSFCell r1c3 = row1.createCell(2);
			r1c3.setCellValue("AGE");
			// Create Second Row
			XSSFRow row2 = sheet.createRow(1);
			XSSFCell r2c1 = row2.createCell(0);
			r2c1.setCellValue("1");
			XSSFCell r2c2 = row2.createCell(1);
			r2c2.setCellValue("Ram");
			XSSFCell r2c3 = row2.createCell(2);
			r2c3.setCellValue("20");
			// Create Third Row
			XSSFRow row3 = sheet.createRow(2);
			XSSFCell r3c1 = row3.createCell(0);
			r3c1.setCellValue("2");
			XSSFCell r3c2 = row3.createCell(1);
			r3c2.setCellValue("Shyam");
			XSSFCell r3c3 = row3.createCell(2);
			r3c3.setCellValue("25");
			workbook.write(fileStream);
			System.out.println("Excel file was created at: " + outputFile);
		} catch (IOException e) {
			System.err.println("Could not create XLSX file at " + outputFile);
			e.printStackTrace();
		}
	}
}