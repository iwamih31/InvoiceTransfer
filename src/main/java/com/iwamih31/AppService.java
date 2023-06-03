package com.iwamih31;

import javax.servlet.http.HttpServletResponse;

import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class AppService {

	public String output_Excel(
			MultipartFile output_model,
			MultipartFile input_model,
			MultipartFile input_data,
			HttpServletResponse httpServletResponse) {
		String message = null;
		Excel excel = new Excel();
		message = excel.transfer_Excel(
				 output_model,
				 input_model,
				 input_data,
				 httpServletResponse);
		return message;
	}

}
