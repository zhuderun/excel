package com.isumi.excel.model;


import java.io.Serializable;

/**
 * @author Administrator
 *
 */
public abstract class ImportModel implements Serializable {

	private static final long serialVersionUID = -2925584890258276687L;
	protected int lineNumber;

	public int getLineNumber() {
		return lineNumber;
	}

	public void setLineNumber(int lineNumber) {
		this.lineNumber = lineNumber;
	}
}
