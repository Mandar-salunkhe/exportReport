package com.deliveryreview.request;

public class DesignationData {
	
	private String id;
	private String ratio;
	private String value;
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	public String getRatio() {
		return ratio;
	}
	public void setRatio(String ratio) {
		this.ratio = ratio;
	}
	public String getValue() {
		return value;
	}
	public void setValue(String value) {
		this.value = value;
	}
	@Override
	public String toString() {
		return "DesignationData [id=" + id + ", ratio=" + ratio + ", value=" + value + "]";
	}
	
}
