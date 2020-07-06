package mx.edev.java.excel.updater;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import mx.edev.java.excel.ExcelExecution;

public class ExcelUpdater implements ExcelExecution {
	
	@Override
	public void exec(String filePath) {

		tryManualClose(filePath);
		tryAutoCloseable(filePath);
		
	}

	private void tryManualClose(String filePath) {
		try {

			FileInputStream fis = new FileInputStream(new File(filePath));
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheetAt(0);
			XSSFRow row1 = sheet.getRow(1);
			XSSFCell cell1 = row1.getCell(1);
			cell1.setCellValue("Mahesh (updated)");
			XSSFRow row2 = sheet.getRow(2);
			XSSFCell cell2 = row2.getCell(1);
			cell2.setCellValue("Ramesh (updated)");
			fis.close();
			FileOutputStream fos = new FileOutputStream(new File(filePath));
			workbook.write(fos);
			fos.close();
			
			System.out.println("Excel file was updated (manual close)");
		} catch (FileNotFoundException e) {
			System.err.println("File not found at [" + filePath + "]");
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void tryAutoCloseable(String filePath) {
		XSSFWorkbook workbook = null;
		try (FileInputStream fis = new FileInputStream(new File(filePath))) {
			workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheetAt(0);
			XSSFRow row1 = sheet.getRow(1);
			XSSFCell cell1 = row1.getCell(2);
			cell1.setCellValue("25 (updated)");
			XSSFRow row2 = sheet.getRow(2);
			XSSFCell cell2 = row2.getCell(2);
			cell2.setCellValue("30 (updated)");
		} catch (FileNotFoundException e) {
			System.err.println("File not found at [" + filePath + "]");
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		try (FileOutputStream fos = new FileOutputStream(new File(filePath))) {
			workbook.write(fos);
			System.out.println("Excel file was updated (AutoClosable)");
		} catch (FileNotFoundException e) {
			System.err.println("File not found at [" + filePath + "]");
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}