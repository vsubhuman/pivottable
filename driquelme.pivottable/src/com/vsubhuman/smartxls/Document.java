package com.vsubhuman.smartxls;

import com.smartxls.WorkBook;

public class Document {

	private DocumentFormat documentFormat;
	private String path;
	private String password;
	
	public Document(DocumentFormat documentFormat, String path) {
		this(documentFormat, path, null);
	}
	
	public Document(DocumentFormat documentFormat, String path, String password) {
		
		if (documentFormat == null || path == null)
			throw new IllegalArgumentException(
					"Document format and path cannot be null!");
		
		this.documentFormat = documentFormat;
		this.path = path;
		this.password = password;
	}

	public DocumentFormat getDocumentFormat() {
		return documentFormat;
	}

	public void setDocumentFormat(DocumentFormat documentFormat) {
		this.documentFormat = documentFormat;
	}

	public String getPath() {
		return path;
	}

	public void setPath(String path) {
		this.path = path;
	}

	public String getPassword() {
		return password;
	}

	public void setPassword(String password) {
		this.password = password;
	}
	
	public WorkBook read() throws Exception {

		return getDocumentFormat().read(getPath(), getPassword());
	}
	
	public void write(WorkBook wb) throws Exception {
		
		getDocumentFormat().write(wb, getPath(), getPassword());
	}
}
