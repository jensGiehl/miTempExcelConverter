package de.agiehl.mitemp.converter.MiTempConverter.service;

import java.io.ByteArrayInputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.util.Assert;

import de.agiehl.mitemp.converter.MiTempConverter.model.MiData;

@Component
public class ReadDataService {

	private static final String SHEET_NAME_ALL = "All";

	private static final int EXPECTED_CELL_SIZE = 6;

	public List<MiData> extractDataFromInputfile(byte[] bytes) throws ReadException {
		try (HSSFWorkbook workbook = new HSSFWorkbook(new ByteArrayInputStream(bytes))) {
			HSSFSheet sheet = workbook.getSheet(SHEET_NAME_ALL);
			Assert.notNull(sheet, () -> "A sheet with name '" + SHEET_NAME_ALL + "' is required!");

			List<MiData> allRows = new ArrayList<>(sheet.getLastRowNum() + 1);
			for (int rowCount = sheet.getFirstRowNum() + 1; rowCount < sheet.getLastRowNum(); rowCount++) {
				allRows.add(extractDataFromRow(sheet, rowCount));
			}

			return allRows;
		} catch (Exception e) {
			throw new ReadException("Error while reading input file", e);
		}
	}

	private MiData extractDataFromRow(HSSFSheet sheet, int rowCount) {
		HSSFRow row = sheet.getRow(rowCount);
		assertCorrectCellCount(rowCount, row);

		return MiData.builder()//
				.name(getName(row))//
				.date(getDate(row))//
				.temperature(getTemperature(row))//
				.humidity(getHumidity(row))//
				.build();
	}

	private Integer getHumidity(HSSFRow row) {
		return Integer.parseInt(row.getCell(3).getStringCellValue());
	}

	private Double getTemperature(HSSFRow row) {
		return Double.parseDouble(row.getCell(2).getStringCellValue());
	}

	private LocalDateTime getDate(HSSFRow row) {
		String dateAsString = row.getCell(1).getStringCellValue();

		return LocalDateTime.parse(dateAsString, DateTimeFormatter.ISO_DATE_TIME);
	}

	private String getName(HSSFRow row) {
		return row.getCell(0).getStringCellValue();
	}

	private void assertCorrectCellCount(int rowCount, HSSFRow row) {
		if (row.getLastCellNum() != EXPECTED_CELL_SIZE) {
			throw new IllegalArgumentException("Error in row " + rowCount + ": Expected " + EXPECTED_CELL_SIZE
					+ " Cells but it was " + row.getLastCellNum());
		}
	}

}
