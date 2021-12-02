package de.agiehl.mitemp.converter.MiTempConverter.service;

import static java.util.stream.Collectors.groupingBy;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.context.i18n.LocaleContextHolder;
import org.springframework.stereotype.Component;

import de.agiehl.mitemp.converter.MiTempConverter.model.MiData;

@Component
public class CreateExcelService {

	@Autowired
	private MessageSource messageSource;

	public byte[] createWorkook(List<MiData> inputData) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFCreationHelper createHelper = workbook.getCreationHelper();
		XSSFCellStyle dateCellStyle = workbook.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(getMessage("output.date.format")));

		inputData.stream()//
				.collect(groupingBy(MiData::getName))//
				.forEach((name, data) -> createSheet(workbook, dateCellStyle, name, data));

		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		workbook.write(outputStream);

		return outputStream.toByteArray();
	}

	private XSSFSheet createSheet(XSSFWorkbook workbook, XSSFCellStyle dateCellStyle, String name, List<MiData> data) {
		XSSFSheet sheet = workbook.createSheet(name);

		int rowCount = 0;
		XSSFRow headline = sheet.createRow(rowCount);
		headline.createCell(0).setCellValue(getMessage("output.headline.date"));
		headline.createCell(1).setCellValue(getMessage("output.headline.temperature"));
		headline.createCell(2).setCellValue(getMessage("output.headline.humidity"));

		for (MiData dataRow : data) {
			XSSFRow row = sheet.createRow(++rowCount);
			XSSFCell dateCell = row.createCell(0);
			dateCell.setCellValue(dataRow.getDate());
			dateCell.setCellStyle(dateCellStyle);

			row.createCell(1).setCellValue(dataRow.getTemperature());
			row.createCell(2).setCellValue(dataRow.getHumidity());
		}

		return sheet;
	}

	private String getMessage(String id) {
		return messageSource.getMessage(id, null, LocaleContextHolder.getLocale());
	}

}
