package org.joerg;

import org.joerg.excel.reader.ExcelFile;
import org.joerg.excel.reader.ExcelFilePoiExample;

public class Start {
	public static void main(String[] args) {
		ExcelFile excelFile = new ExcelFile("Testtabelle.xlsx");

		// List unsupported formulas
		excelFile.listUnsupportedFunctionsOfAllSheets();

		// List cell information of these formula cells
		excelFile.listCellInfoOfUnsupportedFunctionsOfAllSheets();

		// excelFile.setCurrentSheet("Test1");
		// System.out.println("\nCurrent Sheet: " + excelFile.getCurrentSheetName());

		// System.out.println("\nBefore:");
		// System.out.println(excelFile.getCellValue(0, 0));
		// System.out.println(excelFile.getCellValue(1, 0));
		// System.out.println("Result: " + excelFile.getCellValue(2, 0));
		//
		// // Change cell values
		// excelFile.setCellValue(0, 0, 1);
		// excelFile.setCellValue(1, 0, 1);
		//
		// // Run calculation
		// excelFile.evaluateFormulaAt(2, 0);
		// // excelFile.evaluateAllFormulas();
		//
		// System.out.println("\nAfter:");
		// System.out.println(excelFile.getCellValue(0, 0));
		// System.out.println(excelFile.getCellValue(1, 0));
		// System.out.println("Result: " + excelFile.getCellValue(2, 0));

		// From POI Examples
		ExcelFilePoiExample efpe = new ExcelFilePoiExample("Testtabelle.xlsx");
		efpe.listUnsupportedFunctionsOfAllSheets();
	}
}
