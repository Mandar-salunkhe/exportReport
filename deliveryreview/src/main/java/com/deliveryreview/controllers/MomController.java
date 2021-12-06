package com.deliveryreview.controllers;

import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.deliveryreview.request.MomRequest;
import com.deliveryreview.service.MomService;
import com.deliveryreview.service.ServiceResponse;

@RestController
@CrossOrigin(origins = "*", allowedHeaders = "*")
@RequestMapping(value = "/deliveryReviewReport")
public class MomController {

	@Autowired
	MomService momService;

	private static final Logger logger = LogManager.getLogger(MomController.class);
	
	@PostMapping("/downloadMomReport")
	public ServiceResponse exportMomReport(@RequestBody List<MomRequest> momDetails) throws Exception {
		logger.info("START : MOM Report");
		return momService.exportMomReport(momDetails);
	}

}
