package com.openxcell.writer.iface;

import com.openxcell.io.FileHolder;

/**
 * @author vicky.thakor
 * @since 2018-05-15
 */
public interface Writer<T> {
	public void write(FileHolder fileHolder, T data);
}
