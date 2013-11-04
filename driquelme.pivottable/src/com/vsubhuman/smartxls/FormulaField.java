package com.vsubhuman.smartxls;

import com.smartxls.BookPivotField;
import com.smartxls.BookPivotRange;

public class FormulaField extends DataField {

	private static final String TMP_NAME = "_formula";
	
	private String formula;
	
	public FormulaField(String formula) {
		this(formula, null);
	}
	
	public FormulaField(String formula, String name) {
		this(formula, name, null);
	}
	
	public FormulaField(String formula, String name, String numberFormatting) {
		this(formula, name, numberFormatting, null);
	}
	
	public FormulaField(String formula, String name, String numberFormatting, SummarizeType summarizeType) {
		super(null, name, numberFormatting, summarizeType);
		this.formula = formula;
	}
	
	public String getFormula() {
		return formula;
	}

	public void setFormula(String formula) {
		this.formula = formula;
	}
	
	@Override
	public BookPivotField createField(BookPivotRange range) throws Exception {

		String name = getName();
		if (name == null)
			name = TMP_NAME;
		else
			name = name + TMP_NAME;
		
		return range.addFormulaField(name, getFormula());
	}
}
