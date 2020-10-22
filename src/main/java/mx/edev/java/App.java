package mx.edev.java;

import static mx.edev.java.utils.Common.FILE_NAME;
import static mx.edev.java.utils.Common.FILE_NAME_2003;

import mx.edev.java.excel.csv.ExcelToCsv;
import mx.edev.java.excel.csv.ExcelToCsv2003;
import mx.edev.java.excel.reader.ExcelReader;
import mx.edev.java.excel.reader.ExcelReader2003;
import mx.edev.java.excel.updater.ExcelUpdater;
import mx.edev.java.excel.updater.ExcelUpdater2003;
import mx.edev.java.excel.writer.ExcelWriter;
import mx.edev.java.excel.writer.ExcelWriter2003;

public class App {
	private static String strTmp;

	public static void main(String[] args) {

		String tempDirPath = getTempFolderPath();

		System.out.println("\n---------- HSSF ----------\n");
		
		String fileXls = tempDirPath + FILE_NAME_2003;
		new ExcelWriter2003().exec(fileXls);
		new ExcelUpdater2003().exec(fileXls);
		new ExcelReader2003().exec(fileXls);
		new ExcelToCsv2003().exec(fileXls);

		System.out.println("\n---------- XSSF ----------\n");
		
		String fileXlsx = tempDirPath + FILE_NAME;
		new ExcelWriter().exec(fileXlsx);
		new ExcelUpdater().exec(fileXlsx);
		new ExcelReader().exec(fileXlsx);
		new ExcelToCsv().exec(fileXlsx);
	}

	private static String getTempFolderPath() {
		if (strTmp == null) {
			strTmp = System.getProperty("java.io.tmpdir");
			System.out.println("OS current temporary directory: " + strTmp);
			System.out.println("OS Name: " + System.getProperty("os.name"));
			System.out.println("OS Version: " + System.getProperty("os.version"));
			System.out.println();
		}

		return strTmp;
	}
}