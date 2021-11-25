package com.deliveryreview.controllers;

import org.json.JSONObject;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@CrossOrigin(origins = "*", allowedHeaders = "*")
@RequestMapping(value = "/deliveryReviewReport")
public class CurrentDeploymentController {
	
	@PostMapping("/downloadDeploymentReport")
    public String DownloadDeploymentStatus(@RequestBody JSONObject resource) {
		System.out.println("API is been Hit!!!!!! ");
		System.out.println("Resource is "+resource);
		return "Success";
      	
    }
	@GetMapping("/test")
	public String testApi() {
		System.out.println("API is been Hit!!!!!! ");
		return "Hello User";
	}		
}
