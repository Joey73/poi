package org.joerg.excel.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joerg.excel.ext.PoiExtension;

public class ExcelFile {
	private static XSSFWorkbook wb = null;
	private static XSSFSheet currentSheet = null;
	
	public ExcelFile() {
		wb = getExcelWorkbook("Testtabelle.xlsx");
	}

	public ExcelFile(String filePath) {
		wb = getExcelWorkbook(filePath);
	}

	public Collection<String> getUnsupportedFunctionsOfAllSheets(){
		Collection<String> functionList = new ArrayList<String>();
		XSSFSheet sheet;		
		
		int numberOfSheets = wb.getNumberOfSheets();
		for (int i = 0; i < numberOfSheets; i++) {
			sheet = wb.getSheetAt(i);
			functionList.addAll(getUnsupportedFunctionsOfSheet(sheet));
		}
		
		return filterDuplicates(functionList);
	}

	public void listUnsupportedFunctionsOfAllSheets(){
		System.out.println("--- Unsupported Functions of all sheets ---");
		for (String unsupportedFunction : getUnsupportedFunctionsOfAllSheets()) {
			System.out.println(unsupportedFunction);
		}
	}
	
	public Collection<String> filterDuplicates(Collection<String> list){
		Collection<String> resultList = new ArrayList<String>();
		Set<String> tmpSet = new HashSet<String>();
		
		for (String listEntry : list) {
			if(tmpSet.add(listEntry)){
				if(!listEntry.equals("N")){
					resultList.add(listEntry);
				}
			}
		}
		
		return resultList;
	}
	
	public Collection<String> getCellInfoOfUnsupportedFunctionsOfAllSheets(){
		Collection<String> cellInfoList = new ArrayList<String>();
		XSSFSheet sheet;
		int numberOfSheets = wb.getNumberOfSheets();

		for (int i = 0; i < numberOfSheets; i++) {
			sheet = wb.getSheetAt(i);
			for (String unsupportedFunction : getUnsupportedFunctionsOfAllSheets()) {
				cellInfoList.addAll(getCellInfoOfUnsupportedFunctionsOfSheet(sheet, unsupportedFunction));
			}
		}
		return cellInfoList;
	}
	
	public void listCellInfoOfUnsupportedFunctionsOfAllSheets(){
		System.out.println("\n--- Cell info of all unsupported functions ---");
		Collection<String> cellInfoOfUnsupportedFunctionsOfAllSheets = getCellInfoOfUnsupportedFunctionsOfAllSheets();
		
		for (String cellInfo : cellInfoOfUnsupportedFunctionsOfAllSheets) {
			System.out.println(cellInfo);
		}
	}
	
	public Collection<String> getUnsupportedFunctionsOfSheet(XSSFSheet sheet){
		Collection<String> functionList = new ArrayList<String>();
		Collection<String> unsupportedFunctionsOfPoi = PoiExtension.getUnsupportedFunctionsOfPoi();
		Collection<String> allSheetFunctions = getAllSheetFunctions(sheet);

		for (String sheetFunction : allSheetFunctions) {
			for (String unsupportedFunctionOfPoi : unsupportedFunctionsOfPoi) {
				if(sheetFunction.contains(unsupportedFunctionOfPoi)){
					if(!functionList.contains(unsupportedFunctionOfPoi)){
						functionList.add(unsupportedFunctionOfPoi);
						//functionList.add(unsupportedFunctionOfPoi + " - in: " + sheetFunction);
					}
				}
			}
		}
		
		return functionList;
	}

	public Collection<String> getCellInfoOfUnsupportedFunctionsOfSheet(XSSFSheet sheet, String function){
		Collection<String> cellInfo = new ArrayList<String>();

		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell != null) {
					if (Cell.CELL_TYPE_FORMULA == cell.getCellType()) {
						if(cell.getCellFormula().contains(function)){
							cellInfo.add("Formula: " + function + ", Sheet: " + sheet.getSheetName() + ", Row: " + (cell.getRowIndex() + 1) + ", Column: " + (cell.getColumnIndex() + 1) + ", CellContent: " + cell.getCellFormula());
						}
					}
				}
			}
		}
		
		return cellInfo;
	}

	public void listUnsupportedFunctionsOfSheet(XSSFSheet sheet){
		Collection<String> unsupportedFunctionsOfSheet = getUnsupportedFunctionsOfSheet(sheet);
		System.out.println("Unsupported Functions of Sheet:");
		for (String unsupportedFunctionOfSheet : unsupportedFunctionsOfSheet) {
			System.out.println(unsupportedFunctionOfSheet);
		}
	}
	
	public Collection<String> getAllSheetFunctions(XSSFSheet sheet){
		Collection<String> functionList = new ArrayList<String>();
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell != null) {
					if (Cell.CELL_TYPE_FORMULA == cell.getCellType()) {
						functionList.add(cell.getCellFormula());
					}
				}
			}
		}
		return functionList;
	}

	public void listAllSheetFunctions(XSSFSheet sheet) {
		Collection<String> allSheetFunctions = getAllSheetFunctions(sheet);
		for (String function : allSheetFunctions) {
			System.out.println(function);
		}
	}

	public XSSFSheet getSheetAt(int index) {
		return wb.getSheetAt(index);
	}

	public String getCurrentSheetName() {
		return currentSheet.getSheetName();
	}

	public String getStringCellValue(int x, int y) {
		Row row = currentSheet.getRow(x);
		Cell cell = row.getCell(y);
		return cell.getStringCellValue();
	}

	public double getCellValue(int x, int y) {
		Row row = currentSheet.getRow(x);
		Cell cell = row.getCell(y);
		return cell.getNumericCellValue();
	}

	public void setCellValue(int x, int y, double value) {
		Row row = currentSheet.getRow(x);
		Cell cell = row.getCell(y);
		cell.setCellValue(value);
	}

	public void evaluateFormulaAt(int x, int y) {
		Row row = currentSheet.getRow(x);
		Cell cell = row.getCell(y);
		FormulaEvaluator evaluator = wb.getCreationHelper()
				.createFormulaEvaluator();
		if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
			evaluator.evaluateFormulaCell(cell);
		}
	}

	public void evaluateAllFormulas() {
		FormulaEvaluator evaluator = wb.getCreationHelper()
				.createFormulaEvaluator();
		evaluator.evaluateAll();
	}

	private static XSSFWorkbook getExcelWorkbook(String fileName) {
		File file = new File(fileName);

		FileInputStream fis = null;
		try {
			fis = new FileInputStream(file);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}

		XSSFWorkbook wb = null;
		try {
			wb = new XSSFWorkbook(fis);
		} catch (IOException e) {
			e.printStackTrace();
		}
		return wb;
	}

	public XSSFWorkbook getWb() {
		return wb;
	}

	public void setWb(XSSFWorkbook wb) {
		ExcelFile.wb = wb;
	}

	public XSSFSheet getCurrentSheet(){
		return currentSheet;
	}
	
	public void setCurrentSheet(String sheetName) {
		currentSheet = wb.getSheet(sheetName);
	}

	public void setCurrentSheet(int sheetNumber) {
		currentSheet = wb.getSheetAt(sheetNumber);
	}
}
