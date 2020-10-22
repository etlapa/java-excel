package mx.edev.java.excel.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class ExcelReader2003 {

	public void exec(String filePath) {

		System.out.println("\nReading Excel file...");
		System.out.println("\n-----------------------------------------");
		try (FileInputStream fis = new FileInputStream(new File(filePath))) {
			HSSFWorkbook workbook = new HSSFWorkbook(fis);
			HSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<Row> ite = sheet.rowIterator();
			while (ite.hasNext()) {
				Row row = ite.next();
				Iterator<Cell> cite = row.cellIterator();
				while (cite.hasNext()) {
					Cell c = cite.next();
					System.out.print(c.toString() + "  ");
				}
				System.out.println();
			}
		} catch (FileNotFoundException e) {
			System.err.println("File not found at [" + filePath + "]");
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		InputStream ExcelFileToRead = null;
		HSSFWorkbook workbook = null;
		try {
			ExcelFileToRead = new FileInputStream(filePath);

			// Getting the workbook instance for xls file
			workbook = new HSSFWorkbook(ExcelFileToRead);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
}