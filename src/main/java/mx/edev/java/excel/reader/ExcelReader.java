package mx.edev.java.excel.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import mx.edev.java.excel.ExcelExecution;

public class ExcelReader implements ExcelExecution {

	@Override
	public void exec(String filePath) {
		System.out.println("\nReading Excel file...");
		System.out.println("\n-----------------------------------------");
		try (FileInputStream fis = new FileInputStream(new File(filePath))) {
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheetAt(0);
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
	}

}
