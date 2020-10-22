package mx.edev.java.excel.updater;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import mx.edev.java.excel.ExcelExecution;

public class ExcelUpdater2003 implements ExcelExecution {
	
	@Override
	public void exec(String filePath) {

		tryManualClose(filePath);
		tryAutoCloseable(filePath);
		
	}

	private void tryManualClose(String filePath) {
		try {

			FileInputStream fis = new FileInputStream(new File(filePath));
			HSSFWorkbook workbook = new HSSFWorkbook(fis);
			HSSFSheet sheet = workbook.getSheetAt(0);
			HSSFRow row1 = sheet.getRow(1);
			HSSFCell cell1 = row1.getCell(1);
			cell1.setCellValue("Mahesh (updated)");
			HSSFRow row2 = sheet.getRow(2);
			HSSFCell cell2 = row2.getCell(1);
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
		HSSFWorkbook workbook = null;
		try (FileInputStream fis = new FileInputStream(new File(filePath))) {
			workbook = new HSSFWorkbook(fis);
			HSSFSheet sheet = workbook.getSheetAt(0);
			HSSFRow row1 = sheet.getRow(1);
			HSSFCell cell1 = row1.getCell(2);
			cell1.setCellValue("25 (updated)");
			HSSFRow row2 = sheet.getRow(2);
			HSSFCell cell2 = row2.getCell(2);
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