package com.deliveryreview.request;

import java.util.List;

public class MomActionRequest {

	public MomActionRequest() {
		// TODO Auto-generated constructor stub
	}

	private String momID;
	private String taskDetails;
	private String status;
	private String completionDate;
	private long completionDateL;
	private List<String> responsible;
	private String remarks;
	private String oldMOMActionTaskDetails;

	public String getMomID() {
		return momID;
	}

	public void setMomID(String momID) {
		this.momID = momID;
	}

	public String getTaskDetails() {
		return taskDetails;
	}

	public void setTaskDetails(String taskDetails) {
		this.taskDetails = taskDetails;
	}

	public String getStatus() {
		return status;
	}

	public void setStatus(String status) {
		this.status = status;
	}

	public String getCompletionDate() {
		return completionDate;
	}

	public void setCompletionDate(String completionDate) {
		this.completionDate = completionDate;
	}

	public long getCompletionDateL() {
		return completionDateL;
	}

	public void setCompletionDateL(long completionDateL) {
		this.completionDateL = completionDateL;
	}

	public List<String> getResponsible() {
		return responsible;
	}

	public void setResponsible(List<String> responsible) {
		this.responsible = responsible;
	}

	public String getRemarks() {
		return remarks;
	}

	public void setRemarks(String remarks) {
		this.remarks = remarks;
	}

	public String getOldMOMActionTaskDetails() {
		return oldMOMActionTaskDetails;
	}

	public void setOldMOMActionTaskDetails(String oldMOMActionTaskDetails) {
		this.oldMOMActionTaskDetails = oldMOMActionTaskDetails;
	}

	@Override
	public String toString() {
		return "MomActionRequest [momID=" + momID + ", taskDetails=" + taskDetails + ", status=" + status
				+ ", completionDate=" + completionDate + ", completionDateL=" + completionDateL + ", responsible="
				+ responsible + ", remarks=" + remarks + ", oldMOMActionTaskDetails=" + oldMOMActionTaskDetails + "]";
	}

}
