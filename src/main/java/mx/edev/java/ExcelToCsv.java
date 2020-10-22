package mx.edev.java;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//https://dzone.com/articles/the-programmers-way-to-convert-excel-to-csv
public class ExcelToCsv {

	private String inputfilePath = "C:\\tmp\\AP364 BANKCARD Interface - Alnova R2 CIF_v1 2.xlsx";
	private String outputCsvPath = "C:\\tmp\\AP364 BANKCARD Interface - Alnova R2 CIF_v1 2.csv";

	public void init() {

		System.out.println("\nReading Excel file...");
		System.out.println("\n-----------------------------------------");

		DataFormatter formatter = new DataFormatter();

		try (FileInputStream fis = new FileInputStream(new File(inputfilePath));
				PrintStream out = new PrintStream(new FileOutputStream(outputCsvPath), true, "UTF-8")) {
			XSSFWorkbook workbook = new XSSFWorkbook(fis);

			for (Sheet sheet : workbook) {
				for (Row row : sheet) {
					boolean firstCell = true;
					for (Cell cell : row) {
						if (!firstCell)
							out.print(',');
						String text = formatter.formatCellValue(cell);
						out.print(text);
						firstCell = false;
					}
					out.println();
				}
			}

		} catch (FileNotFoundException e) {
			System.err.println("File not found at [" + inputfilePath + "]");
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}