package com.deliveryreview.service;

import java.util.HashMap;
import java.util.Map;

import org.springframework.stereotype.Service;

import com.deliveryreview.request.MomRequest;
import com.deliveryreview.response.CustomResponse;
import com.deliveryreview.utility.ResponseStatus;

@Service
public class MomService {

	public MomService() {
		// TODO Auto-generated constructor stub
	}

	public ServiceResponse exportMomReport(MomRequest momDetails) {

		ServiceResponse serviceResponse = new ServiceResponse();
		Map<Object, Object> responseMap = new HashMap<Object, Object>();
		CustomResponse customResponse = null;
		// SettingM settingsObj = settingsRepoService.getSettingsData();
		// JSONObject settingsJson = new JSONObject(settingsObj.getSettings());
		// responseMap.put("appVersion", settingsJson.get("app_version"));
		customResponse = new CustomResponse(ResponseStatus.SUCCESS.getResponseCode(),
				ResponseStatus.SUCCESS.getResponseMessage());
		responseMap.put("response", customResponse);
		serviceResponse.setServiceResponse(responseMap);
		return serviceResponse;

	}

}
