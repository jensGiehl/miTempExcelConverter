package de.agiehl.mitemp.converter.MiTempConverter.model;

import java.time.LocalDateTime;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class MiData {

	private String name;

	private LocalDateTime date;

	private Double temperature;

	private Integer humidity;

}
