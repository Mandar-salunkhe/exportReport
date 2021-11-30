package com.deliveryreview.controllers;

import java.io.IOException;

import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.deliveryreview.request.CurrentDeploymentRequest;
import com.deliveryreview.service.CurrentDeploymentService;
import com.deliveryreview.service.ServiceResponse;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;

@RestController
@CrossOrigin(origins = "*", allowedHeaders = "*")
@RequestMapping(value = "/deliveryReviewReport")
public class CurrentDeploymentController {
	
	@Autowired
	CurrentDeploymentService deploymentService;
	
	@PostMapping("/downloadDeploymentReport")
    public ServiceResponse DownloadDeploymentStatus(@RequestBody CurrentDeploymentRequest downloadXls) throws IOException {
		System.out.println("API is been Hit!!!!!! ");
		 ObjectMapper mapper = new ObjectMapper();
     
			String activeConsultantsRows = mapper.writeValueAsString(downloadXls.getActiveConsultantsRows());
			JSONArray activeConsultantsArray = new JSONArray(activeConsultantsRows);
			
			String inActiveConsultantsRows = mapper.writeValueAsString(downloadXls.getInActiveConsultantsRows());
			JSONArray inActiveConsultantsRowsArray = new JSONArray(inActiveConsultantsRows);
			
			String partnerEcoSystemRows = mapper.writeValueAsString(downloadXls.getPartnerEcoSystemRows());
			JSONArray partnerEcoSystemRowsArray = new JSONArray(partnerEcoSystemRows);
			
			String inActivePartnerEcoSystemRows = mapper.writeValueAsString(downloadXls.getInActivePartnerEcoSystemRows());
			JSONArray inActivePartnerEcoSystemRowsArray = new JSONArray(inActivePartnerEcoSystemRows);
				
        return deploymentService.exportDeploymentData(downloadXls.getHeaderList(),activeConsultantsArray,inActiveConsultantsRowsArray,partnerEcoSystemRowsArray,inActivePartnerEcoSystemRowsArray);
        
    }
	@GetMapping("/test")
	public String testApi() {
		System.out.println("API is been Hit!!!!!! ");
		return "Hello User";
	}		
}
