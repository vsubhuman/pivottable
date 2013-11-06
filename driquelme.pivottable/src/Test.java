
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

public class Test {

	public static void main(String[] args) throws Exception {
		
		System.out.println("start");
		
		XMLProvider prov = new XMLProvider();
//		PivotTable table = prov.loadConfiguration("table.xml"); 
//		System.out.println("imported");
		
		PivotTable table = createTable();
		System.out.println("table created");

		PivotTableConverter.convert(table);
		System.out.println("table converted");
		
		prov.saveConfiguration("table.xml", table);
		System.out.println("exported");
	}
	
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
}