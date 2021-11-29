package com.deliveryreview.request;

import java.util.List;

import org.json.JSONObject;



public class CurrentDeploymentRequest {
	
	private List <HeaderList> headerList;
	private Object activeConsultantsRows;
	private Object inActiveConsultantsRows;
	private Object inActivePartnerEcoSystemRows;
	private Object partnerEcoSystemRows;

	public List<HeaderList> getHeaderList() {
		return headerList;
	}

	public void setHeaderList(List<HeaderList> headerList) {
		this.headerList = headerList;
	}

	public Object getActiveConsultantsRows() {
		return activeConsultantsRows;
	}

	public void setActiveConsultantsRows(Object activeConsultantsRows) {
		this.activeConsultantsRows = activeConsultantsRows;
	}

	public Object getInActiveConsultantsRows() {
		return inActiveConsultantsRows;
	}

	public void setInActiveConsultantsRows(Object inActiveConsultantsRows) {
		this.inActiveConsultantsRows = inActiveConsultantsRows;
	}

	public Object getInActivePartnerEcoSystemRows() {
		return inActivePartnerEcoSystemRows;
	}

	public void setInActivePartnerEcoSystemRows(Object inActivePartnerEcoSystemRows) {
		this.inActivePartnerEcoSystemRows = inActivePartnerEcoSystemRows;
	}

	public Object getPartnerEcoSystemRows() {
		return partnerEcoSystemRows;
	}

	public void setPartnerEcoSystemRows(Object partnerEcoSystemRows) {
		this.partnerEcoSystemRows = partnerEcoSystemRows;
	}

	@Override
	public String toString() {
		return "CurrentDeploymentRequest [headerList=" + headerList + ", activeConsultantsRows=" + activeConsultantsRows
				+ ", inActiveConsultantsRows=" + inActiveConsultantsRows + ", inActivePartnerEcoSystemRows="
				+ inActivePartnerEcoSystemRows + ", partnerEcoSystemRows=" + partnerEcoSystemRows + "]";
	}	
}
