package com.vsubhuman.smartxls;

import java.util.Collections; 
import java.util.Comparator;
import java.util.List;

import com.smartxls.BookPivotArea;
import com.smartxls.BookPivotField;
import com.smartxls.BookPivotRange;
import com.smartxls.BookPivotRangeModel;
import com.smartxls.WorkBook;
import com.smartxls.enums.PivotBuiltInStyles;
import com.vsubhuman.smartxls.SizeUnit.Size;

/**
 * <p>Class provides functionality to convert specified existing {@link WorkBook}
 * by the rules described in the specified {@link PivotField}.</p>
 * 
 * <p>Also provides functionality to use {@link Document} settings from {@link PivotTable}
 * to automatically read source document and write target document.</p>
 * 
 * @author vsubhuman
 * @version 1.0
 */
public class PivotTableConverter {

	/**
	 * <p>Convert documents using configuration of the specified {@link PivotTable}.
	 * Opens document described as source document of the table, converts it
	 * and writes it into the document described as target document of the table.</p>
	 * 
	 * <p><b>Note:</b> source and target documents required to be set in the table.</p>
	 * 
	 * @param table - {@link PivotTable} to use converting configuration from
	 * @return {@link WorkBook} read from source document and converted by specified configuration
	 * @throws IllegalStateException if source or target document is <code>null</code> 
	 * @throws Exception if read, converting, or write process has failed
	 * @since 1.0
	 */
	public static WorkBook convert(PivotTable table) throws IllegalStateException, Exception {
		
		return convert(table, true);
	}
	
	/**
	 * <p>Convert documents using configuration of the specified {@link PivotTable}.
	 * Opens document described as source document of the table and converts it.
	 * If writeTarget parameter is <code>true</code> - writes converted state
	 * into the document described as target document of the table.</p>
	 * 
	 * <p><b>Note:</b> source document required to be set in the table.
	 * If writeTarget parameter is <code>true</code> - target document
	 * is also required.</p>
	 * 
	 * @param table - {@link PivotTable} to use converting configuration from
	 * @param writeTarget - if <code>true</code> converted state will be saved into
	 * target document from table
	 * @return {@link WorkBook} read from source document and converted by specified configuration
	 * @throws IllegalStateException if source document is <code>null</code> or if writeTarget
	 * parameter is <code>true</code> and target document is <code>null</code>
	 * @throws Exception if read, converting, or write process has failed
	 * @since 1.0
	 */
	public static WorkBook convert(PivotTable table, boolean writeTarget) throws IllegalStateException, Exception {
		
		if (table == null)
			throw new IllegalArgumentException("Pivot table cannot be null!");

		Document sourceDocument = table.getSourceDocument();
		Document targetDocument = table.getTargetDocument();
		
		if (sourceDocument == null)
			throw new IllegalStateException(
					"Cannot convert table without source document settings!");
		
		if (writeTarget && targetDocument == null)
			throw new IllegalStateException(
					"Cannot write target without target document settings!");
		
		WorkBook wb = sourceDocument.read();
		
		convert(wb, table);
		
		if (writeTarget)
			targetDocument.write(wb);
		
		return wb;
	}
	
	/**
	 * <p>Convert specified {@link WorkBook} by configuration described in specified
	 * {@link PivotTable}.</p>
	 * 
	 * @param source - {@link WorkBook} to convert
	 * @param table - {@link PivotTable} to use configuration from
	 * @throws IllegalArgumentException - if either workbook or pivot table is <code>null</code>
	 * @throws Exception - if converting process has failed
	 * @since 1.0
	 */
	public static void convert(WorkBook source, PivotTable table) throws IllegalArgumentException, Exception {

		if (source == null)
			throw new IllegalArgumentException("Source workbook cannot be null!");
		
		if (table == null)
			throw new IllegalArgumentException("Pivot table cannot be null!");

		/*
		 * Source sheet
		 */
		
		int sourceSheet = table.getSourceSheet();
		if (sourceSheet >= 0)
			source.setSheet(sourceSheet);
		else
			sourceSheet = source.getSheet();
		
		/*
		 * Source range
		 */
		
		TableRange sourceRange = table.getSourceRange();
		if (sourceRange == null)
			sourceRange = TableRange.createRange(source);
		
		String range = sourceRange.getRange(source);
		
		/*
		 * Create model
		 */
		
		BookPivotRangeModel pmodel = source.getPivotModel();
		pmodel.setList(range);

		/*
		 * Target cell
		 */
		
		TableCell targetCell = table.getTargetCell();
		
		int targetRow = 0, targetCol = 0;
		if (targetCell != null) {
			
			if (targetCell.isNumbers()) {
				
				targetRow = targetCell.getRow();
				targetCol = targetCell.getCol();
			}
			else {

				String cell = targetCell.getCell(source);
				source.setSelection(cell);
				
				targetRow = source.getActiveRow();
				targetCol = source.getActiveCol();
			}
		}
		
		/*
		 * For page area!
		 */
		targetRow += 2;
		
		pmodel.setLocation(0, targetRow, targetCol);

		/*
		 * Sheet name
		 */
		
		String name = table.getName();
		if (name != null) {
			source.setSheetName(0, name);
		}

		/*
		 * Range parameters
		 */
		
		source.setSelection(targetRow, targetCol, targetRow, targetCol);

		BookPivotRange prange = pmodel.getActivePivotRange();
		prange.setDataOnRow(table.isShowDataColumnsOnRow());
		prange.setShowDrill(table.isShowRowButtons());
		prange.setShowRowGrandTotal(table.isShowTotalRow());
		prange.setShowColGrandTotal(table.isShowTotalCol());
		prange.setShowHeader(table.isShowHeader());
		
		String dataCaption = table.getDataCaption();
		if (dataCaption != null)
			prange.setDataCaption(dataCaption);
		
		PivotBuiltInStyles style = table.getStyle();
		if (style != null)
			prange.setTableStyle(style);
		
		/*
		 * Fields
		 */
		
		List<PivotField> fields = table.getFields();
		Collections.sort(fields, COMPARE_BY_AREA);
		
		PivotArea lastArea = null;
		BookPivotArea parea = null;
		
		int columnFields = -1;
		
		for (PivotField f : fields) {
			
			PivotArea newArea = f.getPivotArea();
			if (newArea == null) {
				
				String sourceField = f.getSource();
				if (sourceField == null)
					sourceField = "unknown";
				
				throw new IllegalStateException(
					"Unidentified pivot area for the field: " + sourceField + "!");
			}
			
			if (newArea != lastArea) {
				
				if (newArea == PivotArea.DATA && columnFields < 0) {
					
					columnFields = PivotArea.COLUMN.getArea(prange).getFieldCount();
				}
				
				lastArea = newArea;
				parea = newArea.getArea(prange);
			}

			int lastFieldCount = parea.getFieldCount();
			BookPivotField newField = f.createField(prange);
			
			if (lastFieldCount == parea.getFieldCount())
				parea.addField(newField);
			
			newField = parea.getField(parea.getFieldCount() - 1);
			f.configureField(newField);
		}
		
		/*
		 * Width
		 */
		
		int column = 0;
		int firstColumnField = -1;
		
		PivotArea lastAddedArea = null;
		boolean compactRowAdded = false;
		
		int dataFields = PivotArea.DATA.getArea(prange).getFieldCount();
		if (dataFields > 0 && table.isShowDataColumnsOnRow())
			dataFields = 0;

		boolean[] widthChanged = new boolean[source.getLastCol() + 1];
		
		for (PivotField f : fields) {
			
			/*
			 * Column index
			 */
			
			PivotArea newArea = f.getPivotArea();
			if (newArea == PivotArea.ROW) {
				
				if (lastAddedArea == PivotArea.ROW && !compactRowAdded) {
					
					column++;
				}
				
				compactRowAdded = f.isFieldCompact();
			}
			else if (newArea == PivotArea.COLUMN) {
				
				if (firstColumnField < 0) {
					
					column++;
					firstColumnField = column;
				}
			}
			else if (newArea == PivotArea.DATA) {

				if (lastAddedArea != newArea) {

					if (firstColumnField < 0) {
						
						column++;
						firstColumnField = column;
					}
					else
						column = firstColumnField;
				}
				else if (dataFields > 0) {
				
					column++;
				}
			}
			
			lastAddedArea = newArea;

			Size width = f.getColumnWidth();
			if (width == null)
				continue;

			int startColumn = column;
			int lastColumn = column;
			
			if (newArea == PivotArea.COLUMN || (newArea == PivotArea.DATA && dataFields > 0))
				lastColumn = source.getLastCol();
			
			while (startColumn <= lastColumn) {
			
				int oldWidth = source.getColWidth(startColumn);
				int newWidth = width.getActualSize();
				if (!widthChanged[startColumn] || newWidth > oldWidth) {
					
					source.setColWidth(startColumn, newWidth);
					widthChanged[startColumn] = true;
				}
				
				if (newArea == PivotArea.DATA && dataFields > 0)
					startColumn += dataFields;
				else
					startColumn++;
			}
		}
	}

	/**
	 * Compares pivot fields by area they meant to be put into.
	 * @since 1.0
	 */
	public static final Comparator<PivotField> COMPARE_BY_AREA =
		new Comparator<PivotField>() {
	
			@Override
			public int compare(PivotField o1, PivotField o2) {
				
				if (o1 == o2)
					return 0;
				
				if (o1 == null)
					return -1;
				
				if (o2 == null)
					return 1;
				
				PivotArea pa1 = o1.getPivotArea();
				PivotArea pa2 = o2.getPivotArea();
				
				if (pa1 == pa2)
					return 0;
				
				if (pa1 == null)
					return -1;

				if (pa2 == null)
					return 1;
				
				return pa1.compareTo(pa2);
			}
		};
}
