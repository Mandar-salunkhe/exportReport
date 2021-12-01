package com.deliveryreview.controllers;

import java.util.List;

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

	@PostMapping("/downloadMomReport")
	public ServiceResponse exportMomReport(@RequestBody List<MomRequest> momDetails) throws Exception {
		return momService.exportMomReport(momDetails);
	}
}
