package com.deliveryreview.utility;

import java.io.IOException;
import java.io.InputStream;

public class ImageUtil {

	private static final String PATH_TO_IMAGE = "";
	
	public static final String CLIENT_LOGO = PATH_TO_IMAGE + "client_logo.png";

	public static InputStream getImage(String image) throws IOException {
		InputStream inputStream = null;
		try {
			inputStream = ImageUtil.class.getClassLoader().getResourceAsStream(image);
			return inputStream;
		} catch (Exception ex) {
			inputStream = ImageUtil.class.getClass().getResourceAsStream(image);
			return inputStream;
		}
	}

}
