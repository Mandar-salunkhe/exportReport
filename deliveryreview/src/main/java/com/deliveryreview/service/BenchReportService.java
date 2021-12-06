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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;
import org.springframework.stereotype.Service;

import com.deliveryreview.excelservice.BenchReportXlsService;
import com.deliveryreview.request.BenchReportRequest;
import com.deliveryreview.response.CustomResponse;
import com.deliveryreview.utility.ResponseStatus;

@Service
public class BenchReportService {

	private static final Logger logger = LogManager.getLogger(BenchReportService.class);

	public ServiceResponse exportBenchReport(BenchReportRequest downloadXls) throws IOException {

		BenchReportXlsService service = new BenchReportXlsService();
		ServiceResponse serviceResponse = new ServiceResponse();
		Map<Object, Object> responseMap = new HashMap<Object, Object>();
		CustomResponse customResponse = null;
		boolean isConsolidateReport = false;
		Workbook wb = new XSSFWorkbook();
		Map<Workbook, File> currentDeployment = service.exportBenchReport(downloadXls, wb, isConsolidateReport);

		File result = new File("");
		if (isConsolidateReport) {

		} else {
			for (Entry<Workbook, File> file : currentDeployment.entrySet()) {
				result = file.getValue();
			}
		}

		String ExcelFileString = encodeFileToBase64Binary(result.getName());

		JSONObject excelData;

		if (!ExcelFileString.isEmpty() && result.exists()) {

			excelData = new JSONObject();
			excelData.put("status", "Success");
			excelData.put("fileName", result.getName());
			excelData.put("filePath", result.getAbsolutePath());
			excelData.put("excelBase64String", ExcelFileString);
			customResponse = new CustomResponse(ResponseStatus.SUCCESS.getResponseCode(),
					ResponseStatus.SUCCESS.getResponseMessage(), excelData.toString());
			result.delete();
			logger.info("SUCCESS END : Bench Report");

		} else {
			excelData = new JSONObject();
			excelData.put("status", "Failed");
			customResponse = new CustomResponse(ResponseStatus.FAILED.getResponseCode(),
					ResponseStatus.FAILED.getResponseMessage(), excelData.toString());
			result.delete();
			logger.error("FAILURE END : Bench Report");

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
