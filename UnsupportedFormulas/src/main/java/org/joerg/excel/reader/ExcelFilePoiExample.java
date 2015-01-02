package org.joerg.excel.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.ss.formula.eval.NotImplementedException;
import org.apache.poi.ss.formula.eval.NotImplementedFunctionException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFilePoiExample {
	private static XSSFWorkbook wb = null;

	public ExcelFilePoiExample() {
		wb = getExcelWorkbook("Testtabelle.xlsx");
	}

	public ExcelFilePoiExample(String filePath) {
		wb = getExcelWorkbook(filePath);
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

	public void listUnsupportedFunctionsOfAllSheets() {
		System.out.println("\n--- Unsupported Formulas from POI examples ---");
		// Fetch all the problems
		List<FormulaEvaluationProblems> problems = new ArrayList<FormulaEvaluationProblems>();
		for (int sn = 0; sn < wb.getNumberOfSheets(); sn++) {
			problems.add(getEvaluationProblems(sn));
		}

		// Produce an overall summary
		Set<String> unsupportedFunctions = new TreeSet<String>();
		for (FormulaEvaluationProblems p : problems) {
			unsupportedFunctions.addAll(p.unsupportedFunctions);
		}

		if (unsupportedFunctions.isEmpty()) {
			System.out.println("There are no unsupported formula functions used");
		} else {
			System.out.println("Unsupported formula functions:");
			for (String function : unsupportedFunctions) {
				System.out.println("  " + function);
			}
			System.out.println("Total unsupported functions = " + unsupportedFunctions.size());
		}

		// Report sheet by sheet
		for (int sn = 0; sn < wb.getNumberOfSheets(); sn++) {
			String sheetName = wb.getSheetName(sn);
			FormulaEvaluationProblems probs = problems.get(sn);

			System.out.println();
			System.out.println("Sheet = " + sheetName);

			if (probs.unevaluatableCells.isEmpty()) {
				System.out.println(" All cells evaluated without error");
			} else {
				for (CellReference cr : probs.unevaluatableCells.keySet()) {
					System.out.println(" " + cr.formatAsString() + " - " + probs.unevaluatableCells.get(cr).toString());
				}
			}
		}
	}

	public FormulaEvaluationProblems getEvaluationProblems(int sheetIndex) {
		return getEvaluationProblems(wb.getSheetAt(sheetIndex));
	}

	public FormulaEvaluationProblems getEvaluationProblems(XSSFSheet sheet) {
		Set<String> unsupportedFunctions = new HashSet<String>();
		Map<CellReference, Exception> unevaluatableCells = new HashMap<CellReference, Exception>();
		XSSFFormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

		for (Row r : sheet) {
			for (Cell c : r) {
				try {
					evaluator.evaluate(c);
				} catch (Exception e) {
					if (e instanceof NotImplementedException && e.getCause() != null) {
						// Has been wrapped with cell details, but we know those
						e = (Exception) e.getCause();
					}

					if (e instanceof NotImplementedFunctionException) {
						NotImplementedFunctionException nie = (NotImplementedFunctionException) e;
						unsupportedFunctions.add(nie.getFunctionName());
					}
					unevaluatableCells.put(new CellReference(c), e);
				}
			}
		}

		return new FormulaEvaluationProblems(unsupportedFunctions, unevaluatableCells);
	}

	public static class FormulaEvaluationProblems {
		/** Which used functions are unsupported by POI at this time */
		public Set<String> unsupportedFunctions;
		/** Which cells had unevaluatable formulas, and why? */
		public Map<CellReference, Exception> unevaluatableCells;

		protected FormulaEvaluationProblems(Set<String> unsupportedFunctions, Map<CellReference, Exception> unevaluatableCells) {
			this.unsupportedFunctions = Collections.unmodifiableSet(unsupportedFunctions);
			this.unevaluatableCells = Collections.unmodifiableMap(unevaluatableCells);
		}
	}
}
