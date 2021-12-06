package com.deliveryreview.request;

import java.util.List;

public class MomReportRequest {

	private List<MomRequest> momDetails;

	public List<MomRequest> getMomDetails() {
		return momDetails;
	}

	public void setMomDetails(List<MomRequest> momDetails) {
		this.momDetails = momDetails;
	}

	@Override
	public String toString() {
		return "MomReportRequest [momDetails=" + momDetails + "]";
	}
	
	

}
