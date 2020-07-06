package mx.edev.java;

import static mx.edev.java.utils.Common.FILE_NAME;

import mx.edev.java.excel.reader.ExcelReader;
import mx.edev.java.excel.updater.ExcelUpdater;
import mx.edev.java.excel.writer.ExcelWriter;

public class App {
	public static void main(String[] args) {
		String filePath = null;
		if (args.length > 0) {
			filePath = args[0];
		} else {
			filePath = getTempFolderPath();
		}

		new ExcelWriter().exec(filePath);
		new ExcelUpdater().exec(filePath);
		new ExcelReader().exec(filePath);

	}

	private static String getTempFolderPath() {
		String strTmp = System.getProperty("java.io.tmpdir");
		System.out.println("OS current temporary directory: " + strTmp);
		System.out.println("OS Name: " + System.getProperty("os.name"));
		System.out.println("OS Version: " + System.getProperty("os.version"));
		System.out.println();
		return strTmp + FILE_NAME;
	}
}