package com.vsubhuman.smartxls;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import com.smartxls.WorkBook;

public enum DocumentFormat {

	CSV {
		
		@Override
		public WorkBook read(String path, String separator) throws Exception {

			WorkBook wb = new WorkBook();
			wb.setCSVSeparator(checkSeparator(separator));
			wb.read(path);
			return wb;
		}
		
		@Override
		public void write(WorkBook wb, String path, String separator) throws Exception {

			char oldSep = wb.getCSVSeparator();
			
			wb.setCSVSeparator(checkSeparator(separator));
			wb.writeCSV(path);
			
			wb.setCSVSeparator(oldSep);
		}
		
		private char checkSeparator(String separator) {
			
			if (separator == null)
				return ';';
			
			if ((separator = separator.trim()).length() != 1)
				throw new IllegalArgumentException("Separator should be 1 character long!");
			
			return separator.charAt(0);
		}
	},
	XLS {
		
		@Override
		public WorkBook read(String path, String password) throws Exception {
			
			WorkBook wb = new WorkBook();
			
			if (password == null)
				wb.read(path);
			else
				wb.read(path, password);
			
			return wb;
		}

		@Override
		public void write(WorkBook wb, String path, String password) throws Exception {
			
			if (password == null)
				wb.write(path);
			else
				wb.write(path, password);
		}
	},
	XLSX {
		
		@Override
		public WorkBook read(String path, String password) throws Exception {
			
			WorkBook wb = new WorkBook();
			
			if (password == null)
				wb.readXLSX(path);
			else
				wb.readXLSX(path, password);
			
			return wb;
		}
		
		@Override
		public void write(WorkBook wb, String path, String password) throws Exception {

			if (password == null)
				wb.writeXLSX(path);
			else
				wb.writeXLSX(path, password);
		}
	},
	XLSB {
		
		@Override
		public WorkBook read(String path, String password) throws Exception {
	
			WorkBook wb = new WorkBook();
			FileInputStream fis = null;
			try {

				fis = new FileInputStream(path);
				
				if (password == null)
					wb.readXLSB(fis);
				else
					wb.readXLSB(fis, password);
				
			} finally {
				
				if (fis != null)
					try {
						fis.close();
					} catch (Exception ignore) {}
			}
			
			return wb;
		}

		@Override
		public void write(WorkBook wb, String path, String password) throws Exception {
			
			FileOutputStream fos = null;
			try {

				fos = new FileOutputStream(path);
				wb.writeXLSB(fos);
				
			} finally {
				
				if (fos != null)
					try {
						fos.close();
					} catch (Exception ignore) {}
			}
		}
	};
	
	public abstract WorkBook read(String path, String password) throws Exception;
	public abstract void write(WorkBook wb, String path, String password) throws Exception;
}
