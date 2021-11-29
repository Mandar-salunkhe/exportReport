package com.deliveryreview.request;

import java.util.List;

public class ActivePartConsList {
	
	private List <ResourceList> partnerEcoSystemRows;

	public List<ResourceList> getPartnerEcoSystemRows() {
		return partnerEcoSystemRows;
	}

	public void setPartnerEcoSystemRows(List<ResourceList> partnerEcoSystemRows) {
		this.partnerEcoSystemRows = partnerEcoSystemRows;
	}

	@Override
	public String toString() {
		return "ActivePartConsList [partnerEcoSystemRows=" + partnerEcoSystemRows + "]";
	}
	
	

}
