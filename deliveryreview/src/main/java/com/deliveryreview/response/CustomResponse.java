package com.deliveryreview.response;

public class CustomResponse {

    private final long responseCode;
    private final String responseMessage;
    private final String momExcelFile;
    

    public CustomResponse(long responseCode, String responseMessage,String momExcelFile) {
	this.responseCode = responseCode;
	this.responseMessage = responseMessage;
	this.momExcelFile = momExcelFile;
    }

    public long getResponseCode() {
	return responseCode;
    }

    public String getResponseMessage() {
	return responseMessage;
    }
    
    public String getMomExcelFile() {
		return momExcelFile;
	}

	@Override
	public String toString() {
		return "CustomResponse [responseCode=" + responseCode + ", responseMessage=" + responseMessage
				+ ", momExcelFile=" + momExcelFile + "]";
	}

}
