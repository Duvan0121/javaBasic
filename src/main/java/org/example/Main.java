package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.awt.Color;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;

public class Main {

    public static void main(String[] args) {
        // Ruta de entrada y salida
        // Ruta del archivo HTML de entrada
        String inputHtmlPath = "C:\\Users\\varga\\OneDrive\\Documentos\\Automatizacion\\cr7Antes.html";
        // Ruta del archivo Excel de salida
        String outputExcelPath = "C:\\Users\\varga\\OneDrive\\Documentos\\Automatizacion\\cr7Aadsdantes.xlsx";

        try {
            // Analizar el archivo HTML con Jsoup
            Document doc = Jsoup.parse(new File(inputHtmlPath), "UTF-8");

            // Crear libro y hoja de Excel
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Sheet1");

            // Crear estilos de fuente
            Font boldFont = workbook.createFont();
            boldFont.setBold(true);
            boldFont.setFontName("Arial");
            boldFont.setFontHeightInPoints((short) 12);

            Font normalFont = workbook.createFont();
            normalFont.setBold(false);
            normalFont.setFontName("Calibri");
            normalFont.setFontHeightInPoints((short) 10);

            // Crear estilo de celda con bordes
            XSSFCellStyle borderStyle = (XSSFCellStyle) workbook.createCellStyle();
            borderStyle.setBorderTop(BorderStyle.THIN);
            borderStyle.setBorderBottom(BorderStyle.THIN);
            borderStyle.setBorderLeft(BorderStyle.THIN);
            borderStyle.setBorderRight(BorderStyle.THIN);

            // Obtener las filas de la tabla HTML
            Elements rows = doc.select("table tr");

            // Iterar por las filas de la tabla
            int rowNum = 0;
            for (Element row : rows) {
                Row excelRow = sheet.createRow(rowNum);
                Elements cells = row.select("td, th");
                int cellNum = 0;

                for (Element cell : cells) {
                    // Crear una celda en la hoja de cálculo de Excel
                    Cell excelCell = excelRow.createCell(cellNum);

                    // Configurar estilo de celda
                    XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
                    if (cell.select("strong").isEmpty()) {
                        cellStyle.setFont(normalFont);
                    } else {
                        cellStyle.setFont(boldFont);
                    }

                    // Configurar color de fondo si se especifica
                    String bgColor = cell.attr("bgcolor");
                    if (!bgColor.isEmpty()) {
                        int[] rgb = hexToRgb(bgColor);
                        XSSFColor color = new XSSFColor(new Color(rgb[0], rgb[1], rgb[2]), null);
                        cellStyle.setFillForegroundColor(color);
                        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    }

                    // Aplicar el estilo de bordes a la celda
                    cellStyle.setBorderTop(BorderStyle.THIN);
                    cellStyle.setBorderBottom(BorderStyle.THIN);
                    cellStyle.setBorderLeft(BorderStyle.THIN);
                    cellStyle.setBorderRight(BorderStyle.THIN);
                    excelCell.setCellStyle(cellStyle);

                    // Establecer el valor de la celda
                    excelCell.setCellValue(cell.text());

                    // Manejar colspan
                    String colspan = cell.attr("colspan");
                    if (!colspan.isEmpty()) {
                        int colspanInt = Integer.parseInt(colspan);
                        // Fusionar celdas
                        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, cellNum, cellNum + colspanInt - 1));
                        cellNum += colspanInt - 1;
                    }

                    cellStyle.setAlignment(HorizontalAlignment.CENTER);

                    // Avanzar al siguiente índice de celda
                    sheet.autoSizeColumn(cellNum);
                    cellNum++;
                }

                // Avanzar al siguiente índice de fila
                rowNum++;
            }

            // Guardar el archivo Excel
            try (FileOutputStream fileOut = new FileOutputStream(outputExcelPath)) {
                workbook.write(fileOut);
            }

            // Cerrar el libro de trabajo
            workbook.close();

            System.out.println("Conversión completada con éxito.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Método para convertir un color hexadecimal a valores RGB
    private static int[] hexToRgb(String hexColor) {
        hexColor = hexColor.startsWith("#") ? hexColor.substring(1) : hexColor;
        int r = Integer.parseInt(hexColor.substring(0, 2), 16);
        int g = Integer.parseInt(hexColor.substring(2, 4), 16);
        int b = Integer.parseInt(hexColor.substring(4, 6), 16);
        return new int[]{r, g, b};
    }
}
