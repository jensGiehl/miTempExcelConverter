package de.agiehl.mitemp.converter.MiTempConverter.service;

import java.io.IOException;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import de.agiehl.mitemp.converter.MiTempConverter.model.MiData;

@Service
public class ConvertService {

	@Autowired
	private ReadDataService readService;

	@Autowired
	private CreateExcelService writeService;

	public byte[] convert(byte[] bytes) throws ConvertException, IOException {
		List<MiData> allData = readService.extractDataFromInputfile(bytes);

		return writeService.createWorkook(allData);
	}

}
