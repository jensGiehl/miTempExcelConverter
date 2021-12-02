package de.agiehl.mitemp.converter.MiTempConverter.controller;

import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import de.agiehl.mitemp.converter.MiTempConverter.service.ConvertException;
import de.agiehl.mitemp.converter.MiTempConverter.service.ConvertService;

@Controller
public class IndexController {

	@Autowired
	ConvertService convertService;

	@GetMapping(path = "/")
	public String getIndexPage() {
		return "index";
	}

	@PostMapping(path = "/upload")
	public @ResponseBody ResponseEntity<Void> uploadFile(@RequestParam("file") MultipartFile file, HttpSession session)
			throws IOException, ConvertException {

		session.setAttribute("file", convertService.convert(file.getBytes()));

		return ResponseEntity.ok().build();
	}

	@GetMapping(path = "/downloadXls", produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	public @ResponseBody byte[] downloadXls(HttpSession session, HttpServletResponse response) {

		String filename = LocalDateTime.now().format(DateTimeFormatter.BASIC_ISO_DATE) + "-miTemp.xlsx";
		response.setHeader("Content-Disposition", "attachment; filename=" + filename);

		return (byte[]) session.getAttribute("file");
	}

}
