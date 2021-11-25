package com.deliveryreview.service;

import java.io.Serializable;
import java.util.Map;

public final class ServiceResponse implements Serializable {

    private static final long serialVersionUID = 1L;
    public Map<Object, Object> serviceResponse;

    public Map<Object, Object> getServiceResponse() {
	return serviceResponse;
    }

    public void setServiceResponse(Map<Object, Object> serviceResponse) {
	this.serviceResponse = serviceResponse;
    }

    public String toString() {
	return serviceResponse.toString();
    }

}
