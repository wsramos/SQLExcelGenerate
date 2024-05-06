package com.local;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.util.Iterator;

/**
 * @author William S. Ramos
 *
 */

public class App {

    private static final Logger logger = LogManager.getLogger(App.class);

    public static void main(String[] args) {
        String excelFilePath = "Definir o caminho do arquivo Excel aqui...";
        String sheetName = "Definir o nome da planilha aqui...";
        FileInputStream inputStream;
        Workbook workbook;

        try {
            logger.info("Iniciando a leitura do arquivo Excel...");
            inputStream = new FileInputStream(new File(excelFilePath));
            workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Planilha '" + sheetName + "' não encontrada no arquivo.");
            }
            Iterator<Row> iterator = sheet.iterator();

            StringBuilder sqlScript = new StringBuilder();

            while (iterator.hasNext()) {
                Row nextRow = iterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();
                String campo1 = "";
                String campo2 = "";
                String campo3 = "";

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if (cell.getColumnIndex() == 2) {
                        campo1 = mapCell(cell, campo1);
                    } else if (cell.getColumnIndex() == 4) {
                        campo2 = mapCell(cell, campo2);
                    } else if (cell.getColumnIndex() == 5) {
                        campo3 = mapCell(cell, campo3);
                    }

                }
                if (campo3.equals("Condição")) {
                    sqlScript.append("UPDATE table SET column = '");
                    sqlScript.append(campo2).append("' WHERE id = ").append(campo1);
                    sqlScript.append(";\n");
                }

            }
            logger.info("Finalizando a leitura do arquivo Excel...");
            logger.info("Iniciando a escrita do arquivo SQL...");
            try (Writer fileWriter = new BufferedWriter(new OutputStreamWriter(
                    new FileOutputStream("output.sql"), StandardCharsets.UTF_8))) {
                fileWriter.write(sqlScript.toString());
            }
            logger.info("Finalizando a escrita do arquivo SQL...");

            workbook.close();
            inputStream.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String mapCell(Cell cell, String v) {
        switch (cell.getCellType()) {
            case STRING:
                v = cell.getStringCellValue();
                break;
            case NUMERIC:
                v = String.valueOf(Long.valueOf((long) cell.getNumericCellValue()));
                break;
            case BOOLEAN:
                v = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                v = cell.getCellFormula();
                break;
            case BLANK:
                v = "";
                break;
            default:
                v = "";
        }
        return v;
    }
}
