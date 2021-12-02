package com.deliveryreview.request;

import java.util.List;

public class ConsolidateReportRequest {

	public ConsolidateReportRequest() {
		// TODO Auto-generated constructor stub
	}

	private CurrentDeploymentRequest currentDeploymentDetails;
	private List<MomRequest> momDetails;

	public CurrentDeploymentRequest getCurrentDeploymentDetails() {
		return currentDeploymentDetails;
	}

	public void setCurrentDeploymentDetails(CurrentDeploymentRequest currentDeploymentDetails) {
		this.currentDeploymentDetails = currentDeploymentDetails;
	}

	public List<MomRequest> getMomDetails() {
		return momDetails;
	}

	public void setMomDetails(List<MomRequest> momDetails) {
		this.momDetails = momDetails;
	}

	@Override
	public String toString() {
		return "ConsolidateReportRequest [currentDeploymentDetails=" + currentDeploymentDetails + ", momDetails="
				+ momDetails + "]";
	}

}
