package com.vsubhuman.smartxls;

import com.smartxls.BookPivotField;
import com.smartxls.BookPivotRange;
import com.vsubhuman.smartxls.SizeUnit.Size;

/**
 * <p>Class represents entity of any type of the pivot table field.</p>
 * 
 * <p>Any pivot table field contains {@link PivotArea} at which if should be put,
 * name of the source data, {@link SortType} to sort data, and width of the column
 * it will be placed into.</p>
 * 
 * <p>If you want to create more specific type of a field, please use special classes:
 * <ul>
 * 	<li>{@link RowField}
 * 	<li>{@link DataField}
 * 	<li>{@link FormulaField}
 * </ul></p>
 * 
 * @author vsubhuman
 * @version 1.0
 */
public class PivotField {

	private PivotArea pivotArea;
	private String source;
	private SortType sortType;
	private Size columnWidth;

	/**
	 * <p>Create new pivot table field to place into specified {@link PivotArea},
	 * with specified source name.</p>
	 * 
	 * <p>Sorting of the field and width of the column will not be changed from default.</p>
	 * 
	 * @param pivotArea - area of a pivot table to place field into
	 * @param source - name of the source data
	 * @since 1.0
	 */
	public PivotField(PivotArea pivotArea, String source) {
		this(pivotArea, source, null);
	}
	
	/**
	 * <p>Create new pivot table field to place into specified {@link PivotArea},
	 * with specified source name, specified column {@link Size}.</p>
	 * 
	 * <p>Sorting of the field will not be changed from default.</p>
	 * 
	 * @param pivotArea - area of a pivot table to place field into
	 * @param source - name of the source data
	 * @param width - width of the column in pixels
	 * @see SizeUnit
	 * @see Size
	 * @since 1.0
	 */
	public PivotField(PivotArea pivotArea, String source, Size width) {
		this(pivotArea, source, width, null);
	}

	/**
	 * Create new pivot table field to place into specified {@link PivotArea},
	 * with specified source name, specified column {@link Size} and specified {@link SortType}.
	 * 
	 * @param pivotArea - area of a pivot table to place field into
	 * @param source - name of the source data
	 * @param width - width of the column
	 * @param sortType - type of the sorting
	 * @see SizeUnit
	 * @see Size
	 * @since 1.0
	 */
	public PivotField(PivotArea pivotArea, String source, Size width, SortType sortType) {
		
		this.pivotArea = pivotArea;
		this.source = source;
		
		this.columnWidth = width;
		this.sortType = sortType;
	}
	
	/**
	 * @return name of the source data
	 * @since 1.0
	 */
	public String getSource() {
		return source;
	}
	
	/**
	 * Sets new name of the source data
	 * @param source - new name of the source data
	 * @since 1.0
	 */
	public void setSource(String source) {
		this.source = source;
	}

	/**
	 * @return {@link SortType} for this field
	 * @since 1.0
	 */
	public SortType getSortType() {
		return sortType;
	}
	
	/**
	 * Sets new {@link SortType} for this field.
	 * If value is <code>null</code> - default sort type
	 * won't be changed.
	 * 
	 * @param sortType - new {@link SortType} for this field
	 * @since 1.0
	 */
	public void setSortType(SortType sortType) {
		this.sortType = sortType;
	}
	
	/**
	 * @return width of the column where this field will be placed
	 * @since 1.0
	 */
	public Size getColumnWidth() {
		return columnWidth;
	}
	
	/**
	 * <p>Sets new width of the column, where this field will be placed.
	 * If value is <code>null</code> - default width won't be changed.</p> 
	 * 
	 * <p><b>Note:</b> if multiple fields placed into one column (like
	 * page area fields, or compact row area fields) then once any of the
	 * fields has change the width of the column, from now on it can be changed
	 * only upwards.</b>
	 * 
	 * @param columnWidth - new width for the field column.
	 * @see SizeUnit
	 * @see Size
	 * @since 1.0
	 */
	public void setColumnWidth(Size columnWidth) {
		this.columnWidth = columnWidth;
	}
	
	/**
	 * <p>Sets new width of the column to the specified number of the
	 * specified width units.</p>
	 * 
	 * <p><b>Note:</b> if unit value is <code>null</code> - all sorts
	 * of bad things and calculation error will happen at converting.</b>
	 * 
	 * @param unit - unit of column width
	 * @param columnWidth - number of units
	 * @see #setColumnWidth(SizeUnit.Size)
	 * @see SizeUnit
	 */
	public void setColumnWidth(SizeUnit unit, int columnWidth) {
		setColumnWidth(new SizeUnit.Size(unit, columnWidth));
	}
	
	/**
	 * Sets new width of the column to the specified number of pixels.
	 * 
	 * @param columnWidth - width of the column in pixels
	 * @see #setColumnWidth(SizeUnit.Size)
	 * @see #setColumnWidth(SizeUnit, int)
	 * @see SizeUnit#PIXEL
	 */
	public void setColumnWidthPx(int columnWidth) {
		setColumnWidth(SizeUnit.PIXEL, columnWidth);
	}
	
	/**
	 * Sets new width of the column to the specified number of points.
	 * 
	 * @param columnWidth - width of the column in points
	 * @see #setColumnWidth(SizeUnit.Size)
	 * @see #setColumnWidth(SizeUnit, int)
	 * @see SizeUnit#POINT
	 */
	public void setColumnWidthPt(int columnWidth) {
		setColumnWidth(SizeUnit.POINT, columnWidth);
	}
	
	/**
	 * @return {@link PivotArea} to place this field into
	 * @since 1.0
	 */
	public PivotArea getPivotArea() {
		return pivotArea;
	}
	
	/**
	 * Sets new {@link PivotArea} to place this field into.
	 * If value is <code>null</code> the error will happen
	 * at the table converting. You should be certain that this value
	 * is set properly before you start converting tables.</p>
	 * 
	 * @param pivotArea - new {@link PivotArea} to place field into
	 * @since 1.0
	 */
	public void setPivotArea(PivotArea pivotArea) {
		this.pivotArea = pivotArea;
	}
	
	/**
	 * This method created for the fields placed into "ROW" pivot area.
	 * Please use special class {@link RowField} to describe that kind
	 * of field. 
	 * 
	 * @return <code>true</code> if this field can considered as compact
	 * by any special parameter settings.
	 * @since 1.0
	 */
	public boolean isFieldCompact() {
		return false;
	}
	
	/**
	 * <p>Method creates new {@link BookPivotField} from specified {@link BookPivotRange}
	 * by description of this instance.</p>
	 * 
	 * <p>This method is called by converting algorithm, to get field described by this
	 * class and instance.</p>
	 * 
	 * <p>This method may put created field into pivot range himself, or he can return
	 * it without putting (than field will be put by converting algorithm).</p>
	 * 
	 * <p>When result field added to the pivot table method {@link #configureField(BookPivotField)}
	 * called with that field as the argument.</p>
	 * 
	 * @param range - {@link BookPivotRange} to create new field
	 * @return {@link BookPivotField} created by this instance
	 * @throws Exception - if process of creating the field has faild
	 * @since 1.0
	 */
	public BookPivotField createField(BookPivotRange range) throws Exception {
		
		String source = getSource();
		if (source == null)
			throw new IllegalStateException(
				"Pivot field source cannot be null!");
		
		return range.getField(source);
	}
	
	/**
	 * <p>Method configures created and added to the table field.</p>
	 * 
	 * <p>This method called by converting algorithm after
	 * method {@link #createField(BookPivotRange)} with {@link BookPivotField}
	 * get as a result of that method and added to the result table.</p>
	 * 
	 * @param field - {@link BookPivotField} to configure
	 * @throws Exception - if configuration process has failed
	 * @since 1.0
	 */
	public void configureField(BookPivotField field) throws Exception {

		SortType sort = getSortType();
		if (sort != null && !field.isSummarizeField()) {
			
			try {
				
				field.setSortType(sort.getType());
				
			} catch (Exception ignore) {}
		}
	}
}