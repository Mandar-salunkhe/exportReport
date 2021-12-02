package com.deliveryreview.controllers;

import java.io.IOException;

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
	
	@PostMapping("/downloadBenchReport")
	public ServiceResponse DownloadBenchReport(@RequestBody BenchReportRequest downloadXls) throws IOException{
		return benchReportServ.exportBenchReport(downloadXls);
	}

}
