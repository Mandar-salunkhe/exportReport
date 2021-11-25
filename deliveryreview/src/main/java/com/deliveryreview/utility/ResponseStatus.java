package com.deliveryreview.utility;

public enum ResponseStatus {

    SUCCESS(200, "Success"), FAILED(500, "Failed"), DUPLICATE(401, "Duplicate");

    private long responseCode;
    private String responseMessage;

    ResponseStatus(long responseCode, String responseMessage) {
	this.responseCode = responseCode;
	this.responseMessage = responseMessage;
    }

    public long getResponseCode() {
	return responseCode;
    }

    public String getResponseMessage() {
	return responseMessage;
    }

}
