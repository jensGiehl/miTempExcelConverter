package de.agiehl.mitemp.converter.MiTempConverter.service;

import static java.util.Comparator.naturalOrder;
import static java.util.stream.Collectors.groupingBy;
import static java.util.stream.Collectors.toList;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.util.Comparator;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
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

		inputData.stream()//
				.collect(groupingBy(MiData::getName))//
				.forEach((name, data) -> createSheet(workbook, name, data));

		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		workbook.write(outputStream);

		return outputStream.toByteArray();
	}

	private XSSFSheet createSheet(XSSFWorkbook workbook, String name, List<MiData> data) {
		XSSFSheet sheet = workbook.createSheet(name);

		int rowCount = 0;
		XSSFRow headline = sheet.createRow(rowCount);
		headline.createCell(0).setCellValue(getMessage("output.headline.date"));
		headline.createCell(1).setCellValue(getMessage("output.headline.temperature"));
		headline.createCell(2).setCellValue(getMessage("output.headline.humidity"));

		List<MiData> sortedList = data.stream().sorted(Comparator.comparing(MiData::getDate, naturalOrder()))
				.collect(toList());

		DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern(getMessage("output.date.format"));

		for (MiData dataRow : sortedList) {
			XSSFRow row = sheet.createRow(++rowCount);
			XSSFCell dateCell = row.createCell(0);
			dateCell.setCellValue(dataRow.getDate().format(dateTimeFormatter));

			row.createCell(1).setCellValue(dataRow.getTemperature());
			row.createCell(2).setCellValue(dataRow.getHumidity());
		}

		sheet.autoSizeColumn(0);
		sheet.autoSizeColumn(1);
		sheet.autoSizeColumn(2);

		return sheet;
	}

	private String getMessage(String id) {
		return messageSource.getMessage(id, null, LocaleContextHolder.getLocale());
	}

}
