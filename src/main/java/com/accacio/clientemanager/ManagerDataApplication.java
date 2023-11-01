package com.accacio.clientemanager;

import java.io.FileInputStream;

import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ManagerDataApplication {

	public static void main(String[] args) {
		SpringApplication.run(ManagerDataApplication.class, args);
		manager();
	}

	public static void manager() {

		File pastaExcel = new File("C:\\Users\\lucas\\Downloads\\bia_manager\\");

		File[] arquivos = pastaExcel.listFiles((dir, name) -> name.endsWith(".xlsx"));

		if (arquivos != null) {
			for (File arquivo : arquivos) {
				try {
					FileInputStream excelFile = new FileInputStream(arquivo);
					Workbook workbook = new XSSFWorkbook(excelFile);
					Sheet sheet = workbook.getSheetAt(0);
					Iterator<Row> iterator = sheet.iterator();

					ObjectMapper objectMapper = new ObjectMapper();
					ObjectNode rootNode = objectMapper.createObjectNode();

					while (iterator.hasNext()) {
						Row currentRow = iterator.next();
						Iterator<Cell> cellIterator = currentRow.iterator();

						ObjectNode jsonObject = objectMapper.createObjectNode();

						while (cellIterator.hasNext()) {
							Cell currentCell = cellIterator.next();
							if (currentCell.getCellType() == CellType.STRING) {
								jsonObject.put("Column" + currentCell.getColumnIndex(),
										currentCell.getStringCellValue());
							} else if (currentCell.getCellType() == CellType.NUMERIC) {
								jsonObject.put("Column" + currentCell.getColumnIndex(),
										currentCell.getNumericCellValue());
							}
						}

						rootNode.set("Row" + currentRow.getRowNum(), jsonObject);
					}

					workbook.close();

					String nomeArquivoJson = arquivo.getName().replace(".xlsx", ".json");
					objectMapper.writeValue(new File(nomeArquivoJson), rootNode);

					System.out.println("JSON salvo com sucesso: " + nomeArquivoJson);

				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
}
