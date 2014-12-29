package org.joerg.excel.ext;

import java.util.Collection;

import org.apache.poi.ss.formula.WorkbookEvaluator;

public class PoiExtension {
	public static Collection<String> getSupportedFunctionsOfPoi(){
		return WorkbookEvaluator.getSupportedFunctionNames();
	}
	
	public static void listSupportedFunctionsOfPoi(){
		for (String supportedFunc : getSupportedFunctionsOfPoi()) {
			System.out.print(supportedFunc + " ");
		}
		System.out.println();
	}

	public static Collection<String> getUnsupportedFunctionsOfPoi(){
		return WorkbookEvaluator.getNotSupportedFunctionNames();
	}

	public static void listUnsupportedFunctionsOfPoi() {
		for (String unsupportedFunc : getUnsupportedFunctionsOfPoi()) {
			System.out.print(unsupportedFunc + " ");
		}
		System.out.println();
	}
}
