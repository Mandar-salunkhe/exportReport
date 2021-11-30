package com.deliveryreview.response;

public class CustomResponse {

    private final long responseCode;
    private final String responseMessage;
    private final String excelFile;
    

    public CustomResponse(long responseCode, String responseMessage,String excelFile) {
	this.responseCode = responseCode;
	this.responseMessage = responseMessage;
	this.excelFile = excelFile;
    }

    public long getResponseCode() {
	return responseCode;
    }

    public String getResponseMessage() {
	return responseMessage;
    }
    
    public String getMomExcelFile() {
		return excelFile;
	}

	@Override
	public String toString() {
		return "CustomResponse [responseCode=" + responseCode + ", responseMessage=" + responseMessage
				+ ", excelFile=" + excelFile + "]";
	}

}
