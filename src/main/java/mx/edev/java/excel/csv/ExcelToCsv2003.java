package mx.edev.java.excel.csv;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

/**
 * Taken from
 * https://dzone.com/articles/the-programmers-way-to-convert-excel-to-csv
 * 
 * @author martin.tlapa.davila
 *
 */
public class ExcelToCsv2003 {

	public void exec(String inputfilePath) {

		String outputCsvPath = inputfilePath.replace("xls", "2003.csv");

		System.out.println("\nReading Excel file...");
		System.out.println("\n-----------------------------------------");

		DataFormatter formatter = new DataFormatter();

		try (FileInputStream fis = new FileInputStream(new File(inputfilePath));
				PrintStream out = new PrintStream(new FileOutputStream(outputCsvPath), true, "UTF-8")) {

			// Using the UTF-8 BOM
			byte[] bom = { (byte) 0xEF, (byte) 0xBB, (byte) 0xBF };
			out.write(bom);

			HSSFWorkbook workbook = new HSSFWorkbook(fis);

			for (int i=0;i<workbook.getNumberOfSheets();i++) {
				for (Row row : workbook.getSheetAt(i)) {
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