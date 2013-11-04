package com.vsubhuman.smartxls;

import com.smartxls.BookPivotField;

public class DataField extends PivotField {

	private String name;
	private String numberFormatting;
	private SummarizeType summarizeType;
	
	public DataField(String source) {
		this(source, null);
	}
	
	public DataField(String source, String name) {
		this(source, name, null);
	}
	
	public DataField(String source, String name, String numberFormatting) {
		this(source, name, numberFormatting, null);
	}
	
	public DataField(String source, String name, String numberFormatting, SummarizeType summarizeType) {
		super(PivotArea.DATA, source);
		this.name = name;
		this.numberFormatting = numberFormatting;
		this.summarizeType = summarizeType;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getNumberFormatting() {
		return numberFormatting;
	}
	
	public void setNumberFormatting(String numberFormatting) {
		this.numberFormatting = numberFormatting;
	}

	public SummarizeType getSummarizeType() {
		return summarizeType;
	}

	public void setSummarizeType(SummarizeType summarizeType) {
		this.summarizeType = summarizeType;
	}
	
	@Override
	public void configureField(BookPivotField field) throws Exception {
		
		super.configureField(field);
		
		SummarizeType sum = getSummarizeType();
		String name = getName();
		
		if ((sum != null || name != null) && field.isSummarizeField()) {

			if (name == null)
				name = field.getName();

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
}