package com.deliveryreview.request;

import java.util.List;

public class ConsolidateReportRequest {

	public ConsolidateReportRequest() {
		// TODO Auto-generated constructor stub
	}

	private CurrentDeploymentRequest currentDeploymentDetails;
	private List<MomRequest> momDetails;
	private BenchReportRequest benchDetails;
	private List<ProjectDeliveryStatRequest> projectDeliveryDetails;

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

	public BenchReportRequest getBenchDetails() {
		return benchDetails;
	}

	public void setBenchDetails(BenchReportRequest benchDetails) {
		this.benchDetails = benchDetails;
	}
	
	public List<ProjectDeliveryStatRequest> getProjectDeliveryDetails() {
		return projectDeliveryDetails;
	}

	public void setProjectDeliveryDetails(List<ProjectDeliveryStatRequest> projectDeliveryDetails) {
		this.projectDeliveryDetails = projectDeliveryDetails;
	}

	@Override
	public String toString() {
		return "ConsolidateReportRequest [currentDeploymentDetails=" + currentDeploymentDetails + ", momDetails="
				+ momDetails + ", benchDetails=" + benchDetails + ", projectDeliveryDetails=" + projectDeliveryDetails
				+ "]";
	}

}
