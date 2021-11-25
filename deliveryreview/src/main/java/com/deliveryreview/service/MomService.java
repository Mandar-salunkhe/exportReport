package com.deliveryreview.service;

import java.util.HashMap;
import java.util.Map;

import org.springframework.stereotype.Service;

import com.deliveryreview.excelservice.MomExcelService;
import com.deliveryreview.request.MomRequest;
import com.deliveryreview.response.CustomResponse;
import com.deliveryreview.utility.ResponseStatus;

@Service
public class MomService {

	public MomService() {
		// TODO Auto-generated constructor stub
	}

	public ServiceResponse exportMomReport(MomRequest momDetails) {
		System.out.println(momDetails);
		ServiceResponse serviceResponse = new ServiceResponse();
		MomExcelService momExcelService = new MomExcelService();
		Map<Object, Object> responseMap = new HashMap<Object, Object>();
		CustomResponse customResponse = null;
		String result = momExcelService.exportMomReport(momDetails);
		if (result.equals("Success")) {
			customResponse = new CustomResponse(ResponseStatus.SUCCESS.getResponseCode(),
					ResponseStatus.SUCCESS.getResponseMessage());
		} else {
			customResponse = new CustomResponse(ResponseStatus.FAILED.getResponseCode(),
					ResponseStatus.FAILED.getResponseMessage());
		}

		responseMap.put("response", customResponse);
		serviceResponse.setServiceResponse(responseMap);
		return serviceResponse;

	}

}
