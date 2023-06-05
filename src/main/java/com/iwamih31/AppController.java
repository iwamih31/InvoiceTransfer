package com.iwamih31;

import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import lombok.AllArgsConstructor;

@Controller
@AllArgsConstructor
@RequestMapping("/InvoiceTransfer")
public class AppController {

	@Autowired
	private AppService service;

	/** RequestMappingのURL */
	public String req() {
		return "/InvoiceTransfer";
	}

	@GetMapping("/")
	public String index(
			Model model) {
		add_View_Data_(model, "index", "トップページ");
		return "view";
	}

	@GetMapping("/Setting")
	public String setting(
			Model model) {
		add_View_Data_(model, "setting", "各種設定");
		return "view";
	}

	@PostMapping("/SetMaterial")
	public String setMaterial(
			Model model) {
		add_View_Data_(model, "setMaterial", "データ移行用ファイル選択");
		return "view";
	}

	@PostMapping("/Output/Excel")
	public String output_Excel(
			@RequestParam("output_model") MultipartFile output_model,
			@RequestParam("input_model") MultipartFile input_model,
			@RequestParam("input_data") MultipartFile input_data,
			HttpServletResponse httpServletResponse,
			RedirectAttributes redirectAttributes) {
		String message = service.output_Excel(output_model, input_model, input_data, httpServletResponse);
		redirectAttributes.addFlashAttribute("message", message);
		___console_Out___("message = " + message);
		return "redirect:" + req() + "/";
	}

	/** view 表示に必要な属性データをモデルに登録 */
	private void add_View_Data_(Model model, String template, String title) {
		model.addAttribute("library", template + "::library");
		model.addAttribute("main", template + "::main");
		model.addAttribute("title", title);
		model.addAttribute("req", req());
		System.out.println("template = " + template);
	}

	/** コンソールに String を出力 */
	public static void ___console_Out___(String message) {
		System.out.println("*");
		System.out.println(message);
		System.out.println("*");
	}

}
