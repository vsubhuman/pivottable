package com.vsubhuman.smartxls;

import com.smartxls.BookPivotField;

public class RowField extends PivotField {

	private boolean subtotalTop = true;
	private boolean outline = true;
	private boolean compact = true;
	
	public RowField(String source) {
		this(source, true);
	}
	
	public RowField(String source, boolean outline) {
		this(source, outline, true);
	}
	
	public RowField(String source, boolean outline, boolean compact) {
		this(source, outline, compact, true);
	}
	
	public RowField(String source, boolean outline, boolean compact, boolean subtotalTop) {
		super(PivotArea.ROW, source);
		this.outline = outline;
		this.compact = compact;
		this.subtotalTop = subtotalTop;
	}

	public boolean isSubtotalTop() {
		return subtotalTop;
	}

	public void setSubtotalTop(boolean subtotalTop) {
		this.subtotalTop = subtotalTop;
	}

	public boolean isOutline() {
		return outline;
	}

	public void setOutline(boolean outline) {
		this.outline = outline;
	}

	public boolean isCompact() {
		return compact;
	}

	public void setCompact(boolean compact) {
		this.compact = compact;
	}

	@Override
	public boolean isFieldCompact() {
		
		return isOutline() && isCompact();
	}

	@Override
	public void configureField(BookPivotField field) throws Exception {
		
		super.configureField(field);
		
		field.setLayout(isCompact(), isOutline(), isSubtotalTop());
	}
}
