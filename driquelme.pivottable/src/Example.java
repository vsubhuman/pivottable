
import com.smartxls.enums.PivotBuiltInStyles;
import com.vsubhuman.smartxls.DataField;
import com.vsubhuman.smartxls.DocumentFormat;
import com.vsubhuman.smartxls.FormulaField;
import com.vsubhuman.smartxls.PivotArea;
import com.vsubhuman.smartxls.PivotField;
import com.vsubhuman.smartxls.PivotTable;
import com.vsubhuman.smartxls.PivotTableConverter;
import com.vsubhuman.smartxls.RowField;
import com.vsubhuman.smartxls.XMLProvider;

public class Example {

	public static void main(String[] args) throws Exception {
		
		PivotTable table = null;
		
		// call one of the methods
		// to create or load table
		
		PivotTableConverter.convert(table);
	}
	
	/*
	 * Simplest way to do it.
	 * Just add this method to your code and when you would need to convert
	 * different table - just change values manually.
	 */
	private static PivotTable createTable() {

		PivotTable table = new PivotTable("PIVOT TABLE");
		
		table.setStyle(PivotBuiltInStyles.PivotStyleMedium4);
		
		table.setSourceDocument(DocumentFormat.XLSX, "table.xlsx");
		table.setTargetDocument(DocumentFormat.XLSX, "table1.xlsx");

		table.addField(new PivotField(PivotArea.PAGE, "Points"));
		
		table.addField(new RowField("Retail")).setColumnWidthPx(250);
		table.addField(new RowField("Product"));
		table.addField(new RowField("Store"));
		
		table.addField(new DataField("07/10/2013", "07/10/2013", "0.0")).setColumnWidthPx(110);
		table.addField(new DataField("14/10/2013", "14/10/2013", "0.0")).setColumnWidthPx(110);
		table.addField(new DataField("Growth", "Growth", "0.0")).setColumnWidthPx(90);
		table.addField(new DataField("Stock", "Stock", "0")).setColumnWidthPx(60);
		table.addField(new DataField("Store With Sales", "Store with sales", "0.0")).setColumnWidthPx(110);
		table.addField(new DataField("Stores", "Store with dis", "0")).setColumnWidthPx(110);
		
		String formula =
				"If("
				+ " ('14/10/2013' + '07/10/2013' + '30/09/2013' + '23/09/2013') / 4 = 0,"
				+ "		 0,"
				+ "		 Stock / ( ('14/10/2013' + '07/10/2013' + '30/09/2013' + '23/09/2013') / 4 )"
				+ ")";
		
		table.addField(new FormulaField(formula, "WOI.", "0.00")).setColumnWidthPx(50);
		
		return table;
	}
	
	/*
	 * More clear way.
	 * 
	 * Add this method to your code and when you need to convert a table - call
	 * it with specified parameters.
	 * 
	 * Pass array of the date fields names (like ["07/10/2013", "14/10/2013"]), formula to calculate
	 * WOI, and paths to source file and target file.
	 */
	private static PivotTable createTable(String[] dateFields, String formula, String source, String target) {
		
		PivotTable table = new PivotTable("PIVOT TABLE");
		
		table.setStyle(PivotBuiltInStyles.PivotStyleMedium4);
		
		
		table.setSourceDocument(DocumentFormat.XLSX, source);
		table.setTargetDocument(DocumentFormat.XLSX, target);
		
		table.addField(new PivotField(PivotArea.PAGE, "Points"));
		
		table.addField(new RowField("Retail")).setColumnWidthPx(250);
		table.addField(new RowField("Product"));
		table.addField(new RowField("Store"));
		
		for (String dateField : dateFields)
			table.addField(new DataField(dateField, dateField, "0.0")).setColumnWidthPx(110);
		
		table.addField(new DataField("Growth", "Growth", "0.0")).setColumnWidthPx(90);
		table.addField(new DataField("Stock", "Stock", "0")).setColumnWidthPx(60);
		table.addField(new DataField("Store With Sales", "Store with sales", "0.0")).setColumnWidthPx(110);
		table.addField(new DataField("Stores", "Store with dis", "0")).setColumnWidthPx(110);
		
		table.addField(new FormulaField(formula, "WOI.", "0.00")).setColumnWidthPx(50);
		
		return table;
	}
	
	/*
	 * Change XML file manually and use this method.
	 * 
	 * Read manual for additional info
	 */
	private static PivotTable loadTable() throws Exception {
		
		return new XMLProvider().loadConfiguration("table.xml");
	}
}