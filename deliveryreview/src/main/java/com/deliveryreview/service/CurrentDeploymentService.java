package com.deliveryreview.service;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.codec.binary.Base64;
import org.apache.commons.io.FileUtils;
import org.json.JSONArray;
import org.json.JSONObject;
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
		File result = service.exportMomReport(headerList,activeConsultantsArray,inActiveConsultantsRowsArray,partnerEcoSystemRowsArray,inActivePartnerEcoSystemRowsArray);
//		if (result.equals("Success")) {
//			customResponse = new CustomResponse(ResponseStatus.SUCCESS.getResponseCode(),
//					ResponseStatus.SUCCESS.getResponseMessage());
//		} else {
//			customResponse = new CustomResponse(ResponseStatus.FAILED.getResponseCode(),
//					ResponseStatus.FAILED.getResponseMessage());
//		}
//
//		responseMap.put("response", customResponse);
//		serviceResponse.setServiceResponse(responseMap);
//		return serviceResponse;
		
String ExcelFileString = encodeFileToBase64Binary(result.getName());
		
		JSONObject excelData;
		
		if (!ExcelFileString.isEmpty() && result.exists()) {
			excelData = new JSONObject();
			excelData.put("status", "Success");
			excelData.put("fileName", result.getName());
			excelData.put("filePath", result.getAbsolutePath());
			excelData.put("excelBase64String", ExcelFileString);		
			customResponse = new CustomResponse(ResponseStatus.SUCCESS.getResponseCode(),
					ResponseStatus.SUCCESS.getResponseMessage(),excelData.toString());
			
		} else {
			excelData = new JSONObject();
			excelData.put("status", "Failed");
				customResponse = new CustomResponse(ResponseStatus.FAILED.getResponseCode(),
					ResponseStatus.FAILED.getResponseMessage(),excelData.toString());
			
		}
		responseMap.put("response", customResponse);
		serviceResponse.setServiceResponse(responseMap);
		return serviceResponse;
	}
	
	private static String encodeFileToBase64Binary(String fileName) throws IOException {
	    File file = new File(fileName);
	    byte[] encoded = Base64.encodeBase64(FileUtils.readFileToByteArray(file));
	    return new String(encoded, StandardCharsets.US_ASCII);
	}

}
