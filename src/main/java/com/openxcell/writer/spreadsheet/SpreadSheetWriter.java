package com.openxcell.writer.spreadsheet;

import java.io.IOException;
import java.util.List;
import java.util.Objects;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.openxcell.io.FileHolder;
import com.openxcell.writer.iface.Writer;

/**
 * @author vicky.thakor
 * @since 2018-05-15
 */
public class SpreadSheetWriter<T> implements Writer<List<T>> {

	private static Logger logger = Logger.getLogger(SpreadSheetWriter.class.getName());
	private SpreadSheetBeanManager spreadSheetBeanManager;

	public SpreadSheetWriter(SpreadSheetTemplate spreadSheetTemplate) {
		Objects.requireNonNull(spreadSheetTemplate, "template can not be null");
		spreadSheetBeanManager = new SpreadSheetBeanManager(spreadSheetTemplate);
	}

	@Override
	public void write(FileHolder fileHolder, List<T> data) {
		Objects.requireNonNull(fileHolder, "file can not be null");
		Objects.requireNonNull(data, "data can not be null");
		
		spreadSheetBeanManager.enableStream(true, 1);
		spreadSheetBeanManager.process(data);
		try {
			spreadSheetBeanManager.closeWorkbook(fileHolder);
		} catch (IOException e) {
			logger.log(Level.SEVERE, e.getMessage(), e);
		}
	}
}
