package com.vsubhuman.smartxls;

import com.smartxls.BookPivotField;

/**
 * <p>Class describes entity of the pivot field that should be placed
 * into "ROW" area. Instance of this class is also a {@link PivotField}
 * with area = {@link PivotArea#ROW}.</p>
 * 
 * <p>Class provides additional functionality to describe specifically
 * a row fields.</p>
 * 
 * @author vsubhuman
 * @since 1.0
 */
public class RowField extends PivotField {

	private boolean subtotalTop = true;
	private boolean outline = true;
	private boolean compact = true;
	
	/**
	 * <p>Create new row field with specified source data, and default
	 * parameters of row field.</p>
	 * 
	 * <p>"Outline", "Compact" and "Subtotal top" values is set to <code>true</code>.</p>
	 * 
	 * @param source - name of the source data
	 * @since 1.0
	 */
	public RowField(String source) {
		this(source, true);
	}
	
	/**
	 * <p>Create new row field with specified source data, and specified
	 * parameter of outline state.</p>
	 * 
	 * <p>"Compact" and "Subtotal top" values is set to <code>true</code>.</p>
	 * 
	 * @param source - name of the source data
	 * @param outline - <code>true</code> if row should be outlined
	 * @since 1.0
	 */
	public RowField(String source, boolean outline) {
		this(source, outline, true);
	}
	
	/**
	 * <p>Create new row field with specified source data, and specified
	 * parameters of outline and compact state.</p>
	 * 
	 * <p>"Subtotal top" value is set to <code>true</code>.</p>
	 * 
	 * 
	 * @param source - name of the source data
	 * @param outline - <code>true</code> if row should be outlined
	 * @param compact - <code>true</code> if row should be
	 * compact (only for outlined rows)
	 * @since 1.0
	 */
	public RowField(String source, boolean outline, boolean compact) {
		this(source, outline, compact, true);
	}
	
	/**
	 * Create new row field with specified source data, and specified
	 * parameters of the row field.
	 * 
	 * @param source - name of the source data
	 * @param outline - <code>true</code> if row should be outlined
	 * @param compact - <code>true</code> if row should be
	 * compact (only for outlined rows)
	 * @param subtotalTop - <code>true</code> if row should show
	 * subtotal values in compact fashion (only for outlined rows)
	 * @since 1.0
	 */
	public RowField(String source, boolean outline, boolean compact, boolean subtotalTop) {
		
		super(PivotArea.ROW, source);
		this.outline = outline;
		this.compact = compact;
		this.subtotalTop = subtotalTop;
	}

	/**
	 * @return whether subtotal values of this row will be shown in the compact
	 * manner.
	 * @since 1.0
	 */
	public boolean isSubtotalTop() {
		return subtotalTop;
	}

	/**
	 * Sets whether subtotal values of this row will be shown in the compact
	 * manner (in the top row itself, and not in the additional "total" row).
	 * Works only for "outlined" rows.
	 * 
	 * @param subtotalTop - if <code>true</code> - subtotal will be compact
	 * @since 1.0
	 */
	public void setSubtotalTop(boolean subtotalTop) {
		this.subtotalTop = subtotalTop;
	}

	/**
	 * @return whether this row will be shown outlined
	 * @since 1.0
	 */
	public boolean isOutline() {
		return outline;
	}

	/**
	 * Sets whether this row will be shown outlined (it will be
	 * separated from child rows). 
	 * 
	 * @param outline if <code>true</code> - row will be outlined
	 * @since 1.0
	 */
	public void setOutline(boolean outline) {
		this.outline = outline;
	}

	/**
	 * @return whether this row will be shown as compact
	 * @since 1.0
	 */
	public boolean isCompact() {
		return compact;
	}

	/**
	 * Sets whether this row will be shown as compact (child
	 * rows will be shown in the same column). Works only
	 * for outlined rows.
	 * 
	 * @param compact - if <code>true</code> - row will be shown as
	 * compact.
	 * @since 1.0
	 */
	public void setCompact(boolean compact) {
		this.compact = compact;
	}

	/**
	 * Returns <code>true</code> if both methods {@link #isOutline()}
	 * and {@link #isCompact()} returns <code>true</code>.
	 * 
	 * @return <code>true</code> if this field is outlined and compact
	 * @since 1.0
	 */
	@Override
	public boolean isFieldCompact() {
		
		return isOutline() && isCompact();
	}

	/**
	 * Calls {@link PivotField#configureField(BookPivotField)} and
	 * configures layout parameters for specified {@link BookPivotField}.
	 * 
	 * @since 1.0
	 */
	@Override
	public void configureField(BookPivotField field) throws Exception {
		
		super.configureField(field);
		
		field.setLayout(isCompact(), isOutline(), isSubtotalTop());
	}

	@Override
	public void setPivotArea(PivotArea pivotArea) {
		
		if (pivotArea != PivotArea.ROW)
			throw new IllegalArgumentException(
				"Cannot set specified pivot area for this kind of field!");
	}
}