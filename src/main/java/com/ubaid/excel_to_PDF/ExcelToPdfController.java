package com.ubaid.excel_to_PDF;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.*;

import jakarta.servlet.http.HttpServletResponse;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;

@RestController
@RequestMapping("/api")
public class ExcelToPdfController {

    private static final int MAX_COLUMNS_PER_ROW = 10;
    private static final float DEFAULT_ROW_HEIGHT = 25f;

    @PostMapping(value = "/convert", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public void convertExcelToPdf(@RequestParam("file") MultipartFile file, HttpServletResponse response) {
        try (Workbook workbook = WorkbookFactory.create(file.getInputStream());
             ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.getSheetAt(0);
            Document pdfDoc = new Document(PageSize.A4.rotate());
            PdfWriter.getInstance(pdfDoc, pdfOutputStream);
            pdfDoc.open();

            Font headerFont = new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD);
            Font normalFont = new Font(Font.FontFamily.HELVETICA, 10);

            for (Row row : sheet) {
                int actualColumnCount = row.getLastCellNum();

                for (int start = 0; start < actualColumnCount; start += MAX_COLUMNS_PER_ROW) {
                    int end = Math.min(start + MAX_COLUMNS_PER_ROW, actualColumnCount);
                    int groupSize = end - start;

                    PdfPTable partialTable = new PdfPTable(groupSize);
                    partialTable.setWidthPercentage(100);

                    // Equal width ratios
                    float[] widths = new float[groupSize];
                    for (int w = 0; w < groupSize; w++) widths[w] = 1f;
                    partialTable.setWidths(widths);

                    for (int i = start; i < end; i++) {
                        Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        String value = getCellValueAsString(cell);

                        PdfPCell pdfCell = new PdfPCell(new Phrase(value, row.getRowNum() == 0 ? headerFont : normalFont));
                        pdfCell.setHorizontalAlignment(Element.ALIGN_CENTER);
                        pdfCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
                        pdfCell.setMinimumHeight(DEFAULT_ROW_HEIGHT);
                        pdfCell.setPadding(5f);
                        partialTable.addCell(pdfCell);
                    }

                    pdfDoc.add(partialTable);
                }
            }

            // Add images from Excel if any
            if (sheet instanceof XSSFSheet xssfSheet) {
                for (POIXMLDocumentPart dr : xssfSheet.getRelations()) {
                    if (dr instanceof XSSFDrawing drawing) {
                        for (XSSFShape shape : drawing.getShapes()) {
                            if (shape instanceof XSSFPicture picture) {
                                XSSFPictureData pictureData = picture.getPictureData();
                                Image image = Image.getInstance(pictureData.getData());
                                image.scaleToFit(400, 300);
                                pdfDoc.newPage();
                                pdfDoc.add(image);
                            }
                        }
                    }
                }
            }

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
        CellType cellType = cell.getCellType();

        switch (cellType) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double d = cell.getNumericCellValue();
                    if (d == Math.floor(d)) {
                        return String.valueOf((long) d);
                    } else {
                        return String.valueOf(d);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                CellValue cellValue = evaluator.evaluate(cell);
                switch (cellValue.getCellType()) {
                    case STRING:
                        return cellValue.getStringValue();
                    case NUMERIC:
                        return String.valueOf(cellValue.getNumberValue());
                    case BOOLEAN:
                        return String.valueOf(cellValue.getBooleanValue());
                    default:
                        return "";
                }
            case BLANK:
                return "";
            default:
                return "";
        }
    }
}
