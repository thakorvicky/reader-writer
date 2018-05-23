package com.openxcell.io;

import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.util.Objects;

/**
 * @author vicky.thakor
 * @since 2018-05-15
 */
public class FileHolder extends File {
	private static final long serialVersionUID = 8398003247119213780L;

	private String remotePath;

	public FileHolder(String pathname) {
		super(pathname);
	}

	public FileHolder(URI uri) {
		super(uri);
	}

	public FileHolder(File parent, String child) {
		super(parent, child);
	}

	public FileHolder(String parent, String child) {
		super(parent, child);
	}

	public String getRemotePath() {
		return remotePath;
	}
	
	public void setRemotePath(String remotePath) {
		this.remotePath = remotePath;
	}

	public static FileHolder createTempFile(String prefix, String suffix) throws IOException {
		return new FileHolder(File.createTempFile(prefix, suffix).getAbsolutePath());
	}

	public String getExtension() {
		String fileName = getName();
		if (Objects.nonNull(fileName)) {
			if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0) {
				return fileName.substring(fileName.lastIndexOf(".") + 1);
			}
		}
		return "";
	}

	@Override
	public String toString() {
		return "FileHolder [remotePath=" + remotePath + ", getPath()=" + getPath() + "]";
	}
}
