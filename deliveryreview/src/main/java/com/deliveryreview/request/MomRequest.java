package com.deliveryreview.request;

import java.util.List;

public class MomRequest {

	public MomRequest() {
		// TODO Auto-generated constructor stub
	}

	private String momID;
	private String date;
	private long dateL;
	private String pointsDiscussed;
	private List<String> participants;
	private List<MomActionRequest> momActionList;

	public String getMomID() {
		return momID;
	}

	public void setMomID(String momID) {
		this.momID = momID;
	}

	public String getDate() {
		return date;
	}

	public void setDate(String date) {
		this.date = date;
	}

	public long getDateL() {
		return dateL;
	}

	public void setDateL(long dateL) {
		this.dateL = dateL;
	}

	public String getPointsDiscussed() {
		return pointsDiscussed;
	}

	public void setPointsDiscussed(String pointsDiscussed) {
		this.pointsDiscussed = pointsDiscussed;
	}

	public List<String> getParticipants() {
		return participants;
	}

	public void setParticipants(List<String> participants) {
		this.participants = participants;
	}

	public List<MomActionRequest> getMomActionList() {
		return momActionList;
	}

	public void setMomActionList(List<MomActionRequest> momActionList) {
		this.momActionList = momActionList;
	}

	@Override
	public String toString() {
		return "MomRequest [momID=" + momID + ", date=" + date + ", dateL=" + dateL + ", pointsDiscussed="
				+ pointsDiscussed + ", participants=" + participants + ", momActionList=" + momActionList + "]";
	}

}
