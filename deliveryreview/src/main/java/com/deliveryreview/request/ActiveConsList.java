package com.deliveryreview.request;

import java.util.List;

public class ActiveConsList {
	
	private List <ResourceList> activeConsultantsRows;

	public List<ResourceList> getActiveConsultantsRows() {
		return activeConsultantsRows;
	}

	public void setActiveConsultantsRows(List<ResourceList> activeConsultantsRows) {
		this.activeConsultantsRows = activeConsultantsRows;
	}

	@Override
	public String toString() {
		return "ActiveConsList [activeConsultantsRows=" + activeConsultantsRows + "]";
	}
	
}
