package com.deliveryreview.service;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.json.JSONArray;
import org.springframework.stereotype.Service;

import com.deliveryreview.excelservice.CurrentDeploymentExcelService;
import com.deliveryreview.request.HeaderList;
import com.deliveryreview.response.CustomResponse;
import com.deliveryreview.utility.ResponseStatus;

@Service
public class CurrentDeploymentService {

	public ServiceResponse exportDeploymentData(List<HeaderList> headerList, JSONArray activeConsultantsArray, JSONArray inActiveConsultantsRowsArray,
			JSONArray partnerEcoSystemRowsArray, JSONArray inActivePartnerEcoSystemRowsArray) throws IOException {
		
		CurrentDeploymentExcelService service = new CurrentDeploymentExcelService();
		ServiceResponse serviceResponse = new ServiceResponse();
		Map<Object, Object> responseMap = new HashMap<Object, Object>();
		CustomResponse customResponse = null;
		String result = service.exportMomReport(headerList,activeConsultantsArray,inActiveConsultantsRowsArray,partnerEcoSystemRowsArray,inActivePartnerEcoSystemRowsArray);
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
