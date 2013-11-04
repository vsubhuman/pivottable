package com.vsubhuman.smartxls;

import java.io.OutputStream;

public interface ConfigurationProvider {

	boolean saveConfiguration(OutputStream os, PivotTable configuration) throws Exception;
	
	PivotTable loadConfiguration(OutputStream os) throws Exception;
}
