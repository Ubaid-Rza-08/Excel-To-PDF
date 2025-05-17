package com.ubaid.excel_to_PDF;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.DocumentException;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.*;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

@RestController
@RequestMapping("/api")
public class ExcelToPdfController {

    @PostMapping(value = "/convert", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public void convertExcelToPdf(@RequestParam("file") MultipartFile file, HttpServletResponse response) {
        try (Workbook workbook = WorkbookFactory.create(file.getInputStream());
             ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.getSheetAt(0);
            Document pdfDoc = new Document();
            PdfWriter.getInstance(pdfDoc, pdfOutputStream);
            pdfDoc.open();

            int columnCount = sheet.getRow(0).getPhysicalNumberOfCells();
            PdfPTable table = new PdfPTable(columnCount);

            for (Row row : sheet) {
                for (int i = 0; i < columnCount; i++) {
                    Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    table.addCell(getCellValueAsString(cell));
                }
            }

            pdfDoc.add(table);
            pdfDoc.close();

            response.setContentType("application/pdf");
            response.setHeader("Content-Disposition", "attachment; filename=converted.pdf");
            response.getOutputStream().write(pdfOutputStream.toByteArray());
            response.getOutputStream().flush();

        } catch (Exception e) {
            throw new RuntimeException("Failed to convert Excel to PDF: " + e.getMessage(), e);
        }
    }

    private String getCellValueAsString(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            default -> "";
        };
    }
}
