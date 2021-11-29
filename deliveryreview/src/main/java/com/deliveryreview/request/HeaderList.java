package com.deliveryreview.request;

public class HeaderList {
	
	private String id;
	private String label;
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	public String getLabel() {
		return label;
	}
	public void setLablel(String label) {
		this.label = label;
	}
	@Override
	public String toString() {
		return "HeaderList [id=" + id + ", lable=" + label + "]";
	}

}
