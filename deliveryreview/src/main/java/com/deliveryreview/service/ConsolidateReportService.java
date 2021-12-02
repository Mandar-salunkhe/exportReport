package com.deliveryreview.service;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.codec.binary.Base64;
import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.stereotype.Service;

import com.deliveryreview.excelservice.CurrentDeploymentExcelService;
import com.deliveryreview.excelservice.MomExcelService;
import com.deliveryreview.request.ConsolidateReportRequest;
import com.deliveryreview.response.CustomResponse;
import com.deliveryreview.utility.ResponseStatus;
import com.fasterxml.jackson.databind.ObjectMapper;

@Service
public class ConsolidateReportService {

	private static final Logger logger = LogManager.getLogger(ConsolidateReportService.class);

	public ServiceResponse downloadConsolidateReport(ConsolidateReportRequest consolidateReportDetails)
			throws IOException {

		ServiceResponse serviceResponse = new ServiceResponse();
		Map<Object, Object> responseMap = new HashMap<Object, Object>();
		CustomResponse customResponse = null;
		boolean isConsolidateReport = true;

		CurrentDeploymentExcelService currentDeploymentExcelService = new CurrentDeploymentExcelService();
		MomExcelService momExcelService = new MomExcelService();

		ObjectMapper mapper = new ObjectMapper();

		String activeConsultantsRows = mapper
				.writeValueAsString(consolidateReportDetails.getCurrentDeploymentDetails().getActiveConsultantsRows());
		JSONArray activeConsultantsArray = new JSONArray(activeConsultantsRows);

		String inActiveConsultantsRows = mapper.writeValueAsString(
				consolidateReportDetails.getCurrentDeploymentDetails().getInActiveConsultantsRows());
		JSONArray inActiveConsultantsRowsArray = new JSONArray(inActiveConsultantsRows);

		String partnerEcoSystemRows = mapper
				.writeValueAsString(consolidateReportDetails.getCurrentDeploymentDetails().getPartnerEcoSystemRows());
		JSONArray partnerEcoSystemRowsArray = new JSONArray(partnerEcoSystemRows);

		String inActivePartnerEcoSystemRows = mapper.writeValueAsString(
				consolidateReportDetails.getCurrentDeploymentDetails().getInActivePartnerEcoSystemRows());
		JSONArray inActivePartnerEcoSystemRowsArray = new JSONArray(inActivePartnerEcoSystemRows);

		Map<Workbook, File> consolidateReport = currentDeploymentExcelService.exportCurrDepReport(
				consolidateReportDetails.getCurrentDeploymentDetails().getHeaderList(), activeConsultantsArray,
				inActiveConsultantsRowsArray, partnerEcoSystemRowsArray, inActivePartnerEcoSystemRowsArray,
				isConsolidateReport);
		Workbook consolidateWb = null;

		for (Entry<Workbook, File> workbook : consolidateReport.entrySet()) {

			consolidateWb = workbook.getKey();
		}

		File momReportFile = momExcelService.exportMomReport(consolidateReportDetails.getMomDetails(), consolidateWb,
				isConsolidateReport);
		String ExcelFileString = encodeFileToBase64Binary(momReportFile.getName());

		JSONObject excelData;

		if (!ExcelFileString.isEmpty() && momReportFile.exists()) {

			excelData = new JSONObject();
			excelData.put("status", "Success");
			excelData.put("fileName", momReportFile.getName());
			excelData.put("filePath", momReportFile.getAbsolutePath());
			excelData.put("excelBase64String", ExcelFileString);

			customResponse = new CustomResponse(ResponseStatus.SUCCESS.getResponseCode(),
					ResponseStatus.SUCCESS.getResponseMessage(), excelData.toString());

			logger.info("Success Response");
		} else {

			excelData = new JSONObject();
			excelData.put("status", "Failed");

			customResponse = new CustomResponse(ResponseStatus.FAILED.getResponseCode(),
					ResponseStatus.FAILED.getResponseMessage(), excelData.toString());

			logger.error("Failed Response");
		}

		responseMap.put("response", customResponse);
		serviceResponse.setServiceResponse(responseMap);
		System.out.println(momReportFile.getName());
		return serviceResponse;

	}

	private static String encodeFileToBase64Binary(String fileName) throws IOException {
		File file = new File(fileName);
		byte[] encoded = Base64.encodeBase64(FileUtils.readFileToByteArray(file));
		return new String(encoded, StandardCharsets.US_ASCII);
	}

}
