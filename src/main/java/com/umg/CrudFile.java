/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.umg;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

/**
 *
 * @author Dario
 */
public class CrudFile {

    public CrudFile() {

    }

    @SuppressWarnings("resource")
    public boolean CreateFile(JSONObject obj, String sheetName, String fileName) {
        boolean result = false;
        try {
            Workbook workbook;
            File file = new File(fileName);
            if (file.exists() && file.length() > 0) {
                workbook = new XSSFWorkbook(new FileInputStream(file));
            } else {
                workbook = new XSSFWorkbook();
            }
            Sheet sheet = workbook.getSheet(sheetName);
            int id = 0; // Inicializar el ID
            if (sheet == null) {
                sheet = workbook.createSheet(sheetName);
                // Crear fila de encabezado
                Row headerRow = sheet.createRow(0);
                int cellIndex = 0;

                // Crear la celda para "ID"
                Cell idCell = headerRow.createCell(cellIndex++);
                idCell.setCellValue("ID");

                // Iterar sobre las claves del objeto JSON
                for (String key : obj.keySet()) {
                    Cell cell = headerRow.createCell(cellIndex++);
                    cell.setCellValue(key);
                }
            } else {
                // Obtener el último ID
                int lastRowNum = sheet.getLastRowNum();
                Row lastRow = sheet.getRow(lastRowNum);
                if (lastRow != null) {
                    Cell idCell = lastRow.getCell(0);
                    if (idCell != null) {
                        id = (int) idCell.getNumericCellValue();
                    }
                }
                id++;
            }

            int lastRowNum = sheet.getLastRowNum();
            Row dataRow = sheet.createRow(lastRowNum + 1);
            int cellIndex = 0;

            // Crear la celda para el ID y establecer su valor
            Cell idCell = dataRow.createCell(cellIndex++);
            idCell.setCellValue(id);

            // Iterar sobre las claves del objeto JSON
            for (String key : obj.keySet()) {
                Cell cell = dataRow.createCell(cellIndex++);
                cell.setCellValue(obj.get(key).toString());
            }
            // Iterar sobre las columnas para redimensionarlas
            for (int i = 0; i < cellIndex; i++) {
                sheet.autoSizeColumn(i);
            }
            // Escribir el archivo Excel
            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
                result = true;
            }
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Error: " + ex.getMessage());
        }
        return result;
    }

}
