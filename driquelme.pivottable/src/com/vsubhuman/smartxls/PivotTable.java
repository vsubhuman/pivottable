package com.vsubhuman.smartxls;

import java.util.ArrayList;
import java.util.List;

import com.smartxls.enums.PivotBuiltInStyles;

/**
 * <p>This class is the central core of "pivot table framework".<br>
 * PivotTable is the entity of <b>configuration</b> of future converting
 * of a simple table into pivot one.</p>
 * 
 * <p>In other words, this configuration describes <b>how</b> one table
 * will be converted into another one, and not the table itself.</p>
 * 
 * <p>This class describes resulting pivot table. It provides functionality to
 * set lots of parameters of the future table. Also it contains information
 * about source table, used to build pivot table.</p>
 * 
 * <p>Example:<pre>
 * PivotTable table = new PivotTable();
 * table.setName("Pivot example");
 * table.setShowHeader(false);
 * table.addField(new PivotField(PivotArea.PAGE, "Pages"));
 * 
 * PivotTable table1 = new PivotTable("Pivot example");
 * 
 * PivotTable table2 = new PivotTable("Pivot example", "A1:D12");
 * 
 * PivotTable table3 = new PivotTable("Pivot example", "A1:D12", 1);
 * </pre>
 * 
 * <p>Described "table" will be placed on a new sheet with the name "Pivot example".
 * The header of the pivot table will not be shown. And it will have 1 field
 * in the "page" area, builded upon column "Pages" from a source table.</p>
 * 
 * <p>Table "table1" gets name of the sheet right in the constructor.</p>
 * <p>Table "table2" the same as "table1" but also will use only range "A1:D12" as source.</p>
 * <p>Table "table3" the same as "table2" but also will use sheet by index 1 as source.</p>
 * 
 * <p>Read also about used in example class {@link PivotField}</p>
 * 
 * <br>
 * 
 * <p>Source and target document also can be specified for pivot table. You can
 * specify format of the document, path to the file and password (optional).
 * Presence of the documents is not necessary for the use of a pivot table.
 * If documents is not specified you will be obligate to use your own instance
 * of the class com.smartxls.WorkBook that you have to create or read and save it
 * manually after converting. If source and target documents are specified process
 * of reading and writing of the WorkBook may be automated.</p>
 * 
 * <p>Read also about document entity: {@link Document}<br>
 * And see methods:<br>
 * {@link PivotTable#setSourceDocument(Document)}</p>
 * {@link PivotTable#setTargetDocument(Document)}<br></p>
 * 
 * @author vsubhuman
 * @version 1.0
 */
public class PivotTable {

	/*
	 * Two fields describes source and target document
	 * of this configuration
	 */
	private Document sourceDocument;
	private Document targetDocument;

	// index of the sheet in the source document
	private int sourceSheet = -1;
	
	// range of the source table
	private TableRange sourceRange;
	
	// name of the result table
	private String name;
	
	// target cell of the result table
	private TableCell targetCell;
	
	// fields of the pivot table
	private List<PivotField> fields;
	
	// style of the result table
	private PivotBuiltInStyles style;
	
	/*
	 * States of the result pivot table
	 */
	private boolean showDataColumnsOnRow;
	private boolean showRowButtons = true;
	private boolean showHeader = true;
	private boolean showTotalRow = true;
	private boolean showTotalCol = true;
	
	// caption of the "data" column
	private String dataCaption;
	
	/**
	 * Create new pivot table.
	 * Currently selected sheet will be selected as source.
	 * All data on the source sheet will be used as source range.
	 * Name of the pivot table sheet will not be changed.
	 * 
	 * @since 1.0
	 */
	public PivotTable() {
		this(null);
	}
	
	/**
	 * Create new pivot table with specified name.
	 * Currently selected sheet will be selected as source.
	 * All data on the source sheet will be used as source range.
	 * 
	 * @param name - name of the sheet of the new pivot table.
	 * @since 1.0
	 */
	public PivotTable(String name) {
		
		this(name, (TableRange) null, -1);
	}

	/**
	 * Create new pivot table with specified name and source range.
	 * Currently selected sheet will be selected as source.
	 * 
	 * @param name - name of the sheet of the new pivot table.
	 * @param sourceRange - range of the source table on the source sheet.
	 * @since 1.0
	 */
	public PivotTable(String name, String sourceRange) {
		
		this(name, sourceRange == null ? null : new TableRange(sourceRange), -1);
	}

	/**
	 * Create new pivot table with specified name, source range and number of source sheet.
	 * 
	 * @param name - name of the sheet of the new pivot table.
	 * @param sourceRange - range of the source table on the source sheet
	 * @param sourceSheet - number of the source sheet.
	 * @since 1.0
	 */
	public PivotTable(String name, TableRange sourceRange, int sourceSheet) {

		this.name = name;
		this.sourceRange = sourceRange;
		this.sourceSheet = sourceSheet;
		this.fields = new ArrayList<PivotField>();
	}
	
	/**
	 * <p>Source sheet is the index of the sheet,
	 * where to seek a source data for the pivot table.</p>
	 * 
	 * @return index of the sheet (starting with 0) with source data.
	 * @since 1.0
	 */
	public int getSourceSheet() {
		return sourceSheet;
	}
	
	/**
	 * <p>Source sheet is the index of the sheet,
	 * where to seek a source data for the pivot table.</p>
	 * 
	 * <p>If value is less than zero - currently
	 * selected sheet will be used as source</p>
	 * 
	 * @param sourceSheet - index of the source sheet (starting with 0)
	 * @since 1.0
	 */
	public void setSourceSheet(int sourceSheet) {
		this.sourceSheet = sourceSheet;
	}
	
	/**
	 * <p>Source range is the range of cells that
	 * will be used as source data for pivot table.</p>
	 * 
	 * @return source range of the table
	 * @since 1.0
	 */
	public TableRange getSourceRange() {
		return sourceRange;
	}
	
	/**
	 * <p>Source range is the range of cells that
	 * will be used as source data for pivot table.</p>
	 * 
	 * <p>Sets source range as pairs of coordinates
	 * for top-left cell of the range and
	 * bottom-right cell of the range.</p>
	 * 
	 * <p>Example: range "0, 0, 1, 1" will start at the cell "0, 0"
	 * and end at the cell "1, 1", and will contains 4 cells.</p>
	 * 
	 * @param row1 - row index of the first cell
	 * @param col1 - column index of the first cell
	 * @param row2 - row index of the second cell
	 * @param col2 - column index of the second cell
	 * @since 1.0
	 */
	public void setSourceRange(int row1, int col1, int row2, int col2) {
		setSourceRange(new TableRange(row1, col1, row2, col2));
	}
	
	/**
	 * <p>Source range is the range of cells that
	 * will be used as source data for pivot table.</p>
	 * 
	 * <p>Sets source range as pair of cell addresses in one string
	 * (Excel style).</p>
	 * 
	 * <p>Example: range "A1:B2" will start at the cell "A1"
	 * and end at the cell "B2", and will contains 4 cells.</p>
	 * 
	 * @param sourceRange - range of the source table as string
	 * @since 1.0
	 */
	public void setSourceRange(String sourceRange) {
		setSourceRange(new TableRange(sourceRange));
	}
	
	/**
	 * <p>Source range is the range of cells that
	 * will be used as source data for pivot table.</p>
	 * 
	 * <p>Sets source range as special object entity.</p>
	 * 
	 * @param sourceRange - range of the source table
	 * @see TableRange
	 * @see #setSourceRange(int, int, int, int)
	 * @see #setSourceRange(String)
	 * @since 1.0
	 */
	public void setSourceRange(TableRange sourceRange) {
		this.sourceRange = sourceRange;
	}

	/**
	 * @return name that will be set to the sheet of the result table
	 * @since 1.0 
	 */
	public String getName() {
		return name;
	}

	/**
	 * Sets name that will be set to the sheet of the result table.
	 * If value is null - name won't be changed.
	 * 
	 * @param name - name of the result table sheet.
	 * @since 1.0
	 */
	public void setName(String name) {
		this.name = name;
	}
	
	/**
	 * <p>Target cell is the cell where top-left corner
	 * of the pivot table will be placed.</p>
	 * 
	 * @return target cell of the result table
	 * @since 1.0
	 */
	public TableCell getTargetCell() {
		
		return targetCell;
	}
	
	/**
	 * <p>Target cell is the cell where top-left corner
	 * of the pivot table will be placed.</p>
	 * 
	 * <p>Sets target cell as pair of coordinates.</p>
	 * 
	 * <p>Example: cell "0, 0" is equivalent of the cell "A1".</p>
	 * 
	 * @param row - row index of the target cell
	 * @param col - column index if the target cell
	 * @since 1.0
	 */
	public void setTargetCell(int row, int col) {
		
		setTargetCell(new TableCell(row, col));
	}
	
	/**
	 * <p>Target cell is the cell where top-left corner
	 * of the pivot table will be placed.</p>
	 * 
	 * <p>Sets target cell as string address of the cell (Excel style).</p>
	 * 
	 * <p>Example: cell "A1".</p>
	 * 
	 * @param cell - address of the target cell
	 * @since 1.0
	 */
	public void setTargetCell(String cell) {
		
		setTargetCell(new TableCell(cell));
	}
	
	/**
	 * <p>Target cell is the cell where top-left corner
	 * of the pivot table will be placed.</p>
	 * 
	 * <p>Sets target cell as special cell entity.</p>
	 * 
	 * @param targetCell - target cell of the result table.
	 * @see TableCell
	 * @see #setTargetCell(int, int)
	 * @see #setTargetCell(String)
	 * @since 1.0
	 */
	public void setTargetCell(TableCell targetCell) {
		
		this.targetCell = targetCell;
	}
	
	/**
	 * <p>Adds specified field to this pivot table.</p>
	 * 
	 * @param field - pivot field to add to this table
	 * @throws NullPointerException if specified field is null
	 * @return specified field
	 * @see PivotField
	 * @see RowField
	 * @see DataField
	 * @see FormulaField
	 * @since 1.0
	 */
	public PivotField addField(PivotField field) throws NullPointerException {
		
		if (field == null)
			throw new IllegalArgumentException(
					"Pivot field cannot be null!");
		
		if (!fields.contains(field))
			fields.add(field);
		
		return field;
	}
	
	/**
	 * @return number of the fields in this pivot field
	 * @since 1.0
	 */
	public int getFieldCount() {
		
		return fields.size();
	}
	
	/**
	 * @param index - index of the field to get
	 * @return field by the specified index
	 * @since 1.0
	 */
	public PivotField getField(int index) {
		
		return fields.get(index);
	}
	
	/**
	 * <p>Method returns list of fields in this pivot table.</p>
	 * <p>No changes in returned list alter original list of fields.
	 * Though field objects is not immutable (if not specified so in documentation),
	 * so changes in that object will create changes in table.</p>
	 * 
	 * @return list of all fields in this pivot table
	 * @since 1.0
	 */
	public List<PivotField> getFields() {

		return new ArrayList<PivotField>(fields);
	}
	
	/**
	 * Removes field from this table by specified index.
	 * 
	 * @param index - index of the field to remove
	 * @return removed field
	 * @since 1.0
	 */
	public PivotField removeField(int index) {
		
		return fields.remove(index);
	}
	
	/**
	 * Removes specified field from this table.
	 * 
	 * @param field - field to remove from this table
	 * @return <code>true</code> if field was removed successfully,
	 * and <code>false</code> otherwise.
	 * @since 1.0
	 */
	public boolean removeField(PivotField field) {
		
		return fields.remove(field);
	}
	
	/**
	 * @return style of the result table.
	 * @see PivotBuiltInStyles
	 * @since 1.0
	 */
	public PivotBuiltInStyles getStyle() {
		return style;
	}

	/**
	 * Sets excel built-in style of the result table.
	 * 
	 * @param style - style of the result table
	 * @see PivotBuiltInStyles
	 * @since 1.0
	 */
	public void setStyle(PivotBuiltInStyles style) {
		this.style = style;
	}

	/**
	 * @return whether data columns will be shown on rows.
	 * @since 1.0
	 */
	public boolean isShowDataColumnsOnRow() {
		return showDataColumnsOnRow;
	}

	/**
	 * <p>Sets whether data columns will be shown on rows.</p>
	 * <p>By default: <code>false</code></p>
	 * 
	 * @param showDataColumnsOnRow - if <code>true</code>, data
	 * columns will be shown as rows.
	 * @since 1.0
	 */
	public void setShowDataColumnsOnRow(boolean showDataColumnsOnRow) {
		this.showDataColumnsOnRow = showDataColumnsOnRow;
	}

	/**
	 * @return whether row buttons should be shown
	 * @since 1.0
	 */
	public boolean isShowRowButtons() {
		return showRowButtons;
	}

	/**
	 * <p>Sets whether row buttons should be shown.</p>
	 * <p>By default: <code>true</code></p>
	 * 
	 * @param showRowButtons - if <code>true</code>, buttons will
	 * be shown on row fields
	 * @since 1.0
	 */
	public void setShowRowButtons(boolean showRowButtons) {
		this.showRowButtons = showRowButtons;
	}

	/**
	 * @return whether pivot table header will be shown
	 * @since 1.0
	 */
	public boolean isShowHeader() {
		return showHeader;
	}

	/**
	 * <p>Sets whether pivot table header will be shown.</p>
	 * <p>By default: <code>true</code>.</p>
	 * 
	 * @param showHeader - if <code>false</code>, pivot table
	 * header will be hidden.
	 * @since 1.0
	 */
	public void setShowHeader(boolean showHeader) {
		this.showHeader = showHeader;
	}

	/**
	 * @return whether row of total results will be shown
	 * @since 1.0
	 */
	public boolean isShowTotalRow() {
		return showTotalRow;
	}

	/**
	 * <p>Sets whether row of total results will be shown.</p>
	 * <p>By default: <code>true</code>.</p> 
	 * 
	 * @param showTotalRow - if <code>false</code>, row of total
	 * results will be hidden
	 * @since 1.0
	 */
	public void setShowTotalRow(boolean showTotalRow) {
		this.showTotalRow = showTotalRow;
	}

	/**
	 * @return whether column of total results will be shown.</p>
	 * @since 1.0
	 */
	public boolean isShowTotalCol() {
		return showTotalCol;
	}

	/**
	 * <p>Sets whether column of total results will be shown.</p>
	 * <p>By default: <code>true</code>.</p>
	 * 
	 * @param showTotalCol - if <code>false</code>, column
	 * of total results will be hidden
	 * @since 1.0
	 */
	public void setShowTotalCol(boolean showTotalCol) {
		this.showTotalCol = showTotalCol;
	}
	
	/**
	 * @return data caption that will be set to data column
	 * of the result table
	 * @since 1.0
	 */
	public String getDataCaption() {
		return dataCaption;
	}

	/**
	 * <p>Sets string caption that will be set to the header
	 * of the data column in the result table. If value
	 * is <code>null</code>, default caption will not be changed.</p>
	 * 
	 * @param dataCaption - new caption of the data header
	 * @since 1.0
	 */
	public void setDataCaption(String dataCaption) {
		this.dataCaption = dataCaption;
	}

	/**
	 * @return document used as a source by this table
	 * @see Document
	 * @since 1.0
	 */
	public Document getSourceDocument() {
		return sourceDocument;
	}

	/**
	 * <p>Sets document that will be used as a source for this table.
	 * By default this value is <code>null</code>.<p>
	 * 
	 * <p>Presence of this document is not necessary for use of a pivot table.
	 * See class comments: {@link PivotTable}.</p>
	 * 
	 * @param sourceDocument - document used as a source for this table
	 * @see Document
	 * @see #setSourceDocument(DocumentFormat, String)
	 * @since 1.0
	 */
	public void setSourceDocument(Document sourceDocument) {
		this.sourceDocument = sourceDocument;
	}
	
	/**
	 * <p>Sets document that will be used as a source for this table.
	 * By default this document is <code>null</code>.<p>
	 * 
	 * <p>Presence of this document is not necessary for use of a pivot table.
	 * See class comments: {@link PivotTable}.</p>
	 * 
	 * @param format - format of the source document
	 * @param path - path to the file of the source document
	 * @see DocumentFormat
	 * @see #setSourceDocument(Document)
	 * @since 1.0
	 */
	public void setSourceDocument(DocumentFormat format, String path) {
		this.sourceDocument = new Document(format, path);
	}
	
	/**
	 * @return document used as a target for this table
	 * @see Document
	 * @since 1.0
	 */
	public Document getTargetDocument() {
		return targetDocument;
	}
	
	/**
	 * <p>Sets document that will be used as a target for this table.
	 * By default this value is <code>null</code>.<p>
	 * 
	 * <p>Presence of this document is not necessary for use of a pivot table.
	 * See class comments: {@link PivotTable}.</p>
	 * 
	 * @param targetDocument - document used as a target for this table
	 * @see Document
	 * @see #setTargetDocument(DocumentFormat, String)
	 * @since 1.0
	 */
	public void setTargetDocument(Document targetDocument) {
		this.targetDocument = targetDocument;
	}
	
	/**
	 * <p>Sets document that will be used as a target for this table.
	 * By default this document is <code>null</code>.<p>
	 * 
	 * <p>Presence of this document is not necessary for use of a pivot table.
	 * See class comments: {@link PivotTable}.</p>
	 * 
	 * @param format - format of the target document
	 * @param path - path to the file of the target document
	 * @see DocumentFormat
	 * @see #setTargetDocument(Document)
	 * @since 1.0
	 */
	public void setTargetDocument(DocumentFormat format, String path) {
		this.targetDocument = new Document(format, path);
	}
}