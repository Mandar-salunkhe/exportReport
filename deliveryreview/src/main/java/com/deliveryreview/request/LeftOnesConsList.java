package com.deliveryreview.request;

import java.util.ArrayList;
import java.util.List;

public class LeftOnesConsList {

	private List <ResourceList> inActiveConsultantsRows;

	public List<ResourceList> getInActiveConsultantsRows() {
		return inActiveConsultantsRows;
	}

	public void setInActiveConsultantsRows(List<ResourceList> inActiveConsultantsRows) {
		this.inActiveConsultantsRows = inActiveConsultantsRows;
	}

	@Override
	public String toString() {
		return "LeftOnesConsList [inActiveConsultantsRows=" + inActiveConsultantsRows + "]";
	}
	
}
