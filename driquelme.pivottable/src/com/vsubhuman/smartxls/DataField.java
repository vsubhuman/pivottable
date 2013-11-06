package com.vsubhuman.smartxls;

import com.smartxls.BookPivotField;

/**
 * <p>Class represents entity of a pivot field that should be added
 * in the "DATA" area of the table. Instance of this class is also
 * a {@link PivotField} with area = {@link PivotArea#DATA}.</p>
 * 
 * <p>Class provides additional functionality to describe specifically
 * a data field.</p>
 * 
 * <p>Excel cannot have source data field and pivot field of the same name,
 * so you should set name of the field different from name of the source data.
 * If name of the field set as the same as name of the data source, name is changed by
 * single dot ('.') character added at the end of the name. Example: if user set
 * name of the source as "data" and name of the field as "data" - name of the
 * field will be changed to "data."</p>
 * 
 * @author vsubhuman
 * @version 1.0
 */
public class DataField extends PivotField {

	private String name;
	private String numberFormatting;
	private SummarizeType summarizeType;
	
	/**
	 * <p>Create new data field with specified source data.</p>
	 * 
	 * <p>Name of the field, number formatting and
	 * {@link SummarizeType} set to default values - <code>null</code>.</p>
	 * 
	 * @param source - name of the source data
	 * @since 1.0
	 */
	public DataField(String source) {
		this(source, null);
	}
	
	/**
	 * <p>Create new data field with specified source data,
	 * and name of the field.</p>
	 * 
	 * <p>Number formatting and {@link SummarizeType} set
	 * to default values - <code>null</code>.</p>
	 * 
	 * @param source - name of the source data
	 * @param name - name of the result field
	 * @since 1.0
	 */
	public DataField(String source, String name) {
		this(source, name, null);
	}
	
	/**
	 * <p>Create new data field with specified source data,
	 * name of the field and number formatting.</p>
	 * 
	 * <p>{@link SummarizeType} set to default value - <code>null</code>.</p>
	 * 
	 * @param source - name of the source data
	 * @param name - name of the result field
	 * @param numberFormatting - number formatting
	 * @since 1.0
	 */
	public DataField(String source, String name, String numberFormatting) {
		this(source, name, numberFormatting, null);
	}
	
	/**
	 * Create new data field with specified source data, name of the field,
	 * number formatting and {@link SummarizeType}.
	 * 
	 * @param source - name of the source data
	 * @param name - name of the result field
	 * @param numberFormatting - number formatting
	 * @param summarizeType - type of total summarization
	 * @since 1.0
	 */
	public DataField(String source, String name, String numberFormatting, SummarizeType summarizeType) {
		
		super(PivotArea.DATA, source);
		this.name = name;
		this.numberFormatting = numberFormatting;
		this.summarizeType = summarizeType;
	}

	/**
	 * @return name that will be set to the result pivot field
	 * @since 1.0
	 */
	public String getName() {
		return name;
	}

	/**
	 * <p>Sets new name that will be set to the result pivot field.
	 * By default pivot field named by source data name and summarize type.
	 * If this value is set to <code>null</code> - default name won't
	 * be changed.</p>
	 * 
	 * <p>Default value: <code>null</code>.</p>
	 * 
	 * <p>Name should be different from the name of the source, or
	 * it will be changed. See class documentation.</p>
	 * 
	 * @param name - new name of the field
	 * @since 1.0 
	 */
	public void setName(String name) {
		this.name = name;
	}

	/**
	 * @return string format used for formatting number values
	 * @since 1.0
	 */
	public String getNumberFormatting() {
		return numberFormatting;
	}
	
	/**
	 * <p>Sets new string format for formatting number values.
	 * Excel format strings used. See excel documentation for
	 * full format description. Simple examples:
	 * <ul>
	 * 	<li>"0" - round value to integer number
	 * 	<li>"0.0" - round value to one number after decimal point
	 * 	<li>"0.00" - round value to two numbers after decimal point
	 * 	<li>"000.00" - round value to two numbers after decimal point
	 * and fill it with zeros to 3 numbers before point. 
	 * </ul></p> 
	 * 
	 * <p>If value is <code>null</code> - default
	 * formatting won't be changed.</p>
	 * 
	 * <p>Default value: <code>null</code>.</p>
	 * 
	 * @param numberFormatting - new format string for number values
	 * @since 1.0
	 */
	public void setNumberFormatting(String numberFormatting) {
		this.numberFormatting = numberFormatting;
	}

	/**
	 * @return type of total value summarization for field data
	 * @since 1.0
	 */
	public SummarizeType getSummarizeType() {
		return summarizeType;
	}

	/**
	 * <p>Sets new {@link SummarizeType} for data of this field.
	 * If value is <code>null</code> - default summarize type won't
	 * be changed.</p>
	 * 
	 * <p>Default value: <code>null</code>.</p>
	 * 
	 * @param summarizeType - new type of summarization
	 * for data of this field
	 * @since 1.0
	 */
	public void setSummarizeType(SummarizeType summarizeType) {
		this.summarizeType = summarizeType;
	}
	
	/**
	 * Calls {@link PivotField#configureField(BookPivotField)} and
	 * configures name of the field, summarize type and number formatting.
	 * 
	 * @since 1.0
	 */
	@Override
	public void configureField(BookPivotField field) throws Exception {
		
		super.configureField(field);
		
		SummarizeType sum = getSummarizeType();
		String name = getName();
		
		if ((sum != null || name != null) && field.isSummarizeField()) {

			if (name == null) {
				
				name = field.getName();
			}
			else if (name.equalsIgnoreCase(getSource())) {
				
				name = name + '.';
			}

			int type;
			if (sum == null)
				type = field.getSummarizeFieldCalcType();
			else
				type = sum.getType();

			field.setSummarizeFieldCalcType(type, name);
		}

		String format = getNumberFormatting();
		if (format != null)
			field.setNumberFormatting(format);
	}
	
	@Override
	public void setPivotArea(PivotArea pivotArea) {
		
		if (pivotArea != PivotArea.DATA)
			throw new IllegalArgumentException(
				"Cannot set specified pivot area for this kind of field!");
	}
}