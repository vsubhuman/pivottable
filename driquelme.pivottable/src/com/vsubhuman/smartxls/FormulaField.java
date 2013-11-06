package com.vsubhuman.smartxls;

import com.smartxls.BookPivotField;
import com.smartxls.BookPivotRange;

/**
 * <p>Class represents entity of the pivot field
 * automatically created from specified formula. This type
 * of the field always placed in the "DATA" area. Instance of
 * this class also is a {@link DataField} with source
 * data = <code>null</code>.</p>
 * 
 * <p>Class provides additional functionality to describe
 * specifically a formula field.</p>
 * 
 * <p><b>Note:</b> formula should be created by names of the <b>source data fields</b>
 * and not a columns or names of the pivot fields. Example: user added
 * two pivot fields "pivot1" with source name "data1" and "pivot2" with
 * source name "data2". Then he want to add formula field, that summarize
 * that two fields. Formula should look like: "data1 + data2" (and not like
 * "pivot1 + pivot2" or "B2 + C2").</p> 
 * 
 * @author vsubhuman
 * @since 1.0
 */
public class FormulaField extends DataField {

	/*
	 * Formula field create new source data field and
	 * pivot table field.
	 * 
	 * Temporary name used to name created data field.
	 * 
	 * Excel cannot have source data field and pivot
	 * field of the same name, so source formula field created
	 * by the name of the field and temporary name.
	 * 
	 * If field name is null - source data name created
	 * by the temporary name and new value of the "noNameCounter"
	 */
	private static final String TMP_NAME = "_formula";
	
	/*
	 * Counter used to name formula fields without valid name
	 */
	private static int noNameCounter;
	
	// formula of this field
	private String formula;
	
	/**
	 * <p>Create new formula field with specified formula.</p>
	 * 
	 * <p>Name of the field, number formatting
	 * and Summarize type are set as <code>null</code>.</p>
	 * 
	 * @param formula - formula of the field
	 * @since 1.0
	 */
	public FormulaField(String formula) {
		this(formula, null);
	}
	
	/**
	 * <p>Create new formula field with specified formula and field name.</p>
	 * 
	 * <p>Number formatting and Summarize type are set as <code>null</code>.</p>
	 * 
	 * @param formula - formula of the field
	 * @param name - name of the field
	 * @since 1.0
	 */
	public FormulaField(String formula, String name) {
		this(formula, name, null);
	}
	
	/**
	 * <p>Create new formula field with specified formula, field name
	 * and number formatting string.</p>
	 * 
	 * <p>Summarize type is set as <code>null</code>.</p>
	 * 
	 * @param formula - formula of the field
	 * @param name - name of the field
	 * @param numberFormatting - string format for number values
	 * @since 1.0
	 */
	public FormulaField(String formula, String name, String numberFormatting) {
		this(formula, name, numberFormatting, null);
	}
	
	/**
	 * Create new formula field with specified formula, field name,
	 * number formatting string and {@link SummarizeType}. 
	 * 
	 * @param formula - formula of the field
	 * @param name - name of the field
	 * @param numberFormatting - string format for number values
	 * @param summarizeType - {@link SummarizeType} for the field data
	 * @since 1.0
	 */
	public FormulaField(String formula, String name, String numberFormatting, SummarizeType summarizeType) {
		super(null, name, numberFormatting, summarizeType);
		this.formula = formula;
	}
	
	/**
	 * @return formula of this field
	 * @since 1.0
	 */
	public String getFormula() {
		return formula;
	}

	/**
	 * Sets new formula for this field. Formula is created by
	 * pivot table rules, see class documentation.
	 * 
	 * @param formula - new formula string
	 * @since 1.0
	 */
	public void setFormula(String formula) {
		this.formula = formula;
	}
	
	/**
	 * Creates new formula field and adds it to the pivot range
	 * (by specifics of the SmartXLS)
	 * 
	 * @since 1.0
	 */
	@Override
	public BookPivotField createField(BookPivotRange range) throws Exception {

		String name = getName();
		if (name == null)
			name = ++noNameCounter + TMP_NAME;
		else
			name = name + TMP_NAME;
		
		return range.addFormulaField(name, getFormula());
	}
}
