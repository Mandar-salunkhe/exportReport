package com.deliveryreview.response;

public class CustomResponse {

    private final long responseCode;
    private final String responseMessage;

    public CustomResponse(long responseCode, String responseMessage) {
	this.responseCode = responseCode;
	this.responseMessage = responseMessage;
    }

    public long getResponseCode() {
	return responseCode;
    }

    public String getResponseMessage() {
	return responseMessage;
    }

    @Override
    public String toString() {
	return "CustomResponse [responseCode=" + responseCode + ", responseMessage=" + responseMessage + "]";
    }

}
