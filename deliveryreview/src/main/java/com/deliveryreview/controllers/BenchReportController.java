package com.deliveryreview.controllers;

import java.io.IOException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.deliveryreview.request.BenchReportRequest;
import com.deliveryreview.service.BenchReportService;
import com.deliveryreview.service.ServiceResponse;

@RestController
@CrossOrigin(origins = "*", allowedHeaders = "*")
@RequestMapping(value = "/deliveryReviewReport")
public class BenchReportController {
	
	@Autowired
	BenchReportService benchReportServ;

	private static final Logger logger = LogManager.getLogger(BenchReportController.class);
	
	@PostMapping("/downloadBenchReport")
	public ServiceResponse DownloadBenchReport(@RequestBody BenchReportRequest downloadXls) throws IOException{
		logger.info("START : Bench Report");
		return benchReportServ.exportBenchReport(downloadXls);
	}

}
