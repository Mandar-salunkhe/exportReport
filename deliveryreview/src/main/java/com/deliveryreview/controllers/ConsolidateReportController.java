package com.deliveryreview.controllers;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.deliveryreview.request.ConsolidateReportRequest;
import com.deliveryreview.service.ConsolidateReportService;
import com.deliveryreview.service.CurrentDeploymentService;
import com.deliveryreview.service.MomService;
import com.deliveryreview.service.ServiceResponse;

@RestController
@CrossOrigin(origins = "*", allowedHeaders = "*")
@RequestMapping(value = "/deliveryReviewReport")
public class ConsolidateReportController {

	@Autowired
	MomService momService;

	@Autowired
	CurrentDeploymentService currentDeploymentService;

	@Autowired
	ConsolidateReportService consolidateReportService;
	
	private static final Logger logger = LogManager.getLogger(ConsolidateReportController.class);

	@PostMapping("/downloadConsolidateReport")
	public ServiceResponse downloadConsolidateReport(@RequestBody ConsolidateReportRequest consolidateReportDetails)
			throws Exception {
		logger.info("START : Consolidate Report");
		
		return consolidateReportService.downloadConsolidateReport(consolidateReportDetails);
	}

}
