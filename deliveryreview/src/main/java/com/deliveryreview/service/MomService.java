package com.deliveryreview.service;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.codec.binary.Base64;
import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import org.springframework.stereotype.Service;

import com.deliveryreview.excelservice.MomExcelService;
import com.deliveryreview.request.MomRequest;
import com.deliveryreview.response.CustomResponse;
import com.deliveryreview.utility.ResponseStatus;

@Service
public class MomService {

	
	private static final Logger logger = LogManager.getLogger(MomService.class);
	
	public ServiceResponse exportMomReport(List<MomRequest> momDetails) throws IOException {
		
		ServiceResponse serviceResponse = new ServiceResponse();
		MomExcelService momExcelService = new MomExcelService();
		Map<Object, Object> responseMap = new HashMap<Object, Object>();
		CustomResponse customResponse = null;
		Workbook wb =   new XSSFWorkbook();
		File result = momExcelService.exportMomReport(momDetails,wb,false);
		
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
			result.delete();
			
			logger.info("SUCCESS END : MOM Report");
		} else {
			
			excelData = new JSONObject();
			excelData.put("status", "Failed");
			result.delete();
			customResponse = new CustomResponse(ResponseStatus.FAILED.getResponseCode(),
					ResponseStatus.FAILED.getResponseMessage(),excelData.toString());
			
			logger.error("FAILURE END : MOM Report");
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
