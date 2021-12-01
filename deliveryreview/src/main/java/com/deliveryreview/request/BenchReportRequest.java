package com.deliveryreview.request;

import java.util.ArrayList;
import java.util.List;

public class BenchReportRequest {
	
	private List <String> monthList;
	private ArrayList<DesignationData> aptaList;
	private ArrayList<DesignationData> ptaList;
	private ArrayList<DesignationData> staList;
	private ArrayList<DesignationData> taList;
	private ArrayList<DesignationData> stsList;
	private ArrayList<DesignationData> tsList;
	private ArrayList<DesignationData> sseList;
	private ArrayList<DesignationData> seList;
	private ArrayList<DesignationData> totalList;
	private ArrayList<DesignationData> simpleAverageList;
	
	public List<String> getMonthList() {
		return monthList;
	}
	public void setMonthList(List<String> monthList) {
		this.monthList = monthList;
	}
	public ArrayList<DesignationData> getAptaList() {
		return aptaList;
	}
	public void setAptaList(ArrayList<DesignationData> aptaList) {
		this.aptaList = aptaList;
	}
	public ArrayList<DesignationData> getPtaList() {
		return ptaList;
	}
	public void setPtaList(ArrayList<DesignationData> ptaList) {
		this.ptaList = ptaList;
	}
	public ArrayList<DesignationData> getStaList() {
		return staList;
	}
	public void setStaList(ArrayList<DesignationData> staList) {
		this.staList = staList;
	}
	public ArrayList<DesignationData> getTaList() {
		return taList;
	}
	public void setTaList(ArrayList<DesignationData> taList) {
		this.taList = taList;
	}
	public ArrayList<DesignationData> getStsList() {
		return stsList;
	}
	public void setStsList(ArrayList<DesignationData> stsList) {
		this.stsList = stsList;
	}
	public ArrayList<DesignationData> getTsList() {
		return tsList;
	}
	public void setTsList(ArrayList<DesignationData> tsList) {
		this.tsList = tsList;
	}
	public ArrayList<DesignationData> getSseList() {
		return sseList;
	}
	public void setSseList(ArrayList<DesignationData> sseList) {
		this.sseList = sseList;
	}
	public ArrayList<DesignationData> getSeList() {
		return seList;
	}
	public void setSeList(ArrayList<DesignationData> seList) {
		this.seList = seList;
	}
	public ArrayList<DesignationData> getTotalList() {
		return totalList;
	}
	public void setTotalList(ArrayList<DesignationData> totalList) {
		this.totalList = totalList;
	}
	public ArrayList<DesignationData> getSimpleAverageList() {
		return simpleAverageList;
	}
	public void setSimpleAverageList(ArrayList<DesignationData> simpleAverageList) {
		this.simpleAverageList = simpleAverageList;
	}
	@Override
	public String toString() {
		return "BenchReportRequest [monthList=" + monthList + ", aptaList=" + aptaList + ", ptaList=" + ptaList
				+ ", staList=" + staList + ", taList=" + taList + ", stsList=" + stsList + ", tsList=" + tsList
				+ ", sseList=" + sseList + ", seList=" + seList + ", totalList=" + totalList + ", simpleAverageList="
				+ simpleAverageList + "]";
	}

}
