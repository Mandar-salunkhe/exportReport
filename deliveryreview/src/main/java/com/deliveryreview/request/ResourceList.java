package com.deliveryreview.request;

public class ResourceList {

	private String id;
	private String label;
	//private String title;
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	public String getLabel() {
		return label;
	}
	public void setLabel(String label) {
		this.label = label;
	}

	/*
	 * public String getTitle() { return title; } public void setTitle(String title)
	 * { this.title = title; }
	 */
	@Override
	public String toString() {
		return "ResourceList [id=" + id + ", label=" + label + "]";
	}
	
}
