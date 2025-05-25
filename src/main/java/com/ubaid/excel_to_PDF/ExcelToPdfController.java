package com.ubaid.excel_to_PDF;

import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.colors.Color;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.io.exceptions.IOException;
import com.itextpdf.io.font.constants.StandardFonts;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.Set;
import java.util.HashSet;

@org.springframework.web.bind.annotation.RestController
public class ExcelToPdfController {

    private static final Logger logger = LoggerFactory.getLogger(ExcelToPdfController.class);
    private static final float A4_WIDTH = PageSize.A4.getWidth();
    private static final float A4_HEIGHT = PageSize.A4.getHeight();
    private static final float MARGIN = 36f;
    private static final float MAX_TABLE_WIDTH = A4_WIDTH - 2 * MARGIN;
    private static final int MAX_COLUMNS_PER_TABLE = 5; // For normal tables
    private static final int MAX_COLUMNS = 7; // For calendars
    private static final Color HEADER_BG_COLOR = new DeviceRgb(230, 230, 230);
    private static final Color ROW_EVEN_BG_COLOR = new DeviceRgb(245, 245, 245);
    private static final Color ROW_ODD_BG_COLOR = new DeviceRgb(255, 255, 255);
    private static final Pattern DAY_PATTERN = Pattern.compile("^(Sun|Mon|Tue|Wed|Thu|Fri|Sat|Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday)$", Pattern.CASE_INSENSITIVE);
    private static final Pattern MONTH_PATTERN = Pattern.compile("^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|January|February|March|April|May|June|July|August|September|October|November|December)", Pattern.CASE_INSENSITIVE);
    private static final String[] DAY_HEADERS = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"};
    private static final String[] MONTHS_WITH_NOTES = {"April", "July", "December"};
    private static final int DAYS_IN_WEEK = 7;
    private static final int MAX_WEEKS = 6;

    @PostMapping(value = "/convert-excel-to-pdf", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<Resource> convertExcelToPdf(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) {
            throw new IllegalArgumentException("Uploaded file is empty");
        }

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream());
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            PdfWriter writer = new PdfWriter(baos);
            PdfDocument pdf = new PdfDocument(writer);
            pdf.setDefaultPageSize(PageSize.A4);
            Document document = new Document(pdf, PageSize.A4);
            PdfFont font;
            PdfFont fontBold;
            try {
                font = PdfFontFactory.createFont(StandardFonts.HELVETICA);
                fontBold = PdfFontFactory.createFont(StandardFonts.HELVETICA_BOLD);
            } catch (java.io.IOException e) {
                throw new RuntimeException("Failed to create PDF fonts: " + e.getMessage(), e);
            }

            boolean isCalendarDocument = false;
            if (workbook.getNumberOfSheets() > 0) {
                Sheet firstSheet = workbook.getSheetAt(0);
                Row headerRow = firstSheet.getRow(0);
                if (headerRow != null) {
                    int dayCount = 0;
                    for (int col = 0; col < 7; col++) {
                        String cellValue = getCellValue(headerRow.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK));
                        if (DAY_PATTERN.matcher(cellValue).matches()) {
                            dayCount++;
                        }
                    }
                    if (dayCount >= 5) {
                        isCalendarDocument = true;
                    }
                }
            }

            if (isCalendarDocument) {
                document.add(new Paragraph("2023")
                        .setFont(fontBold)
                        .setFontSize(24)
                        .setTextAlignment(TextAlignment.CENTER)
                        .setMarginTop(PageSize.A4.getHeight() / 2 - 24));
                document.getPdfDocument().addNewPage();
                logger.debug("Added title page with '2023'");
            }

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if (i > 0 || isCalendarDocument) {
                    document.getPdfDocument().addNewPage();
                    logger.debug("Added new page for sheet {}", workbook.getSheetAt(i).getSheetName());
                }
                Sheet sheet = workbook.getSheetAt(i);
                PageSize pageSize = determinePageSize(sheet);
                document.getPdfDocument().setDefaultPageSize(pageSize);
                document.setMargins(MARGIN, MARGIN, MARGIN, MARGIN);
                createSheetTable(document, sheet, font, fontBold);

                List<byte[]> images = extractImagesFromSheet(sheet);
                if (!images.isEmpty()) {
                    for (byte[] imageData : images) {
                        try {
                            com.itextpdf.layout.element.Image pdfImage = new com.itextpdf.layout.element.Image(ImageDataFactory.create(imageData));
                            pdfImage.setAutoScale(true);
                            document.add(pdfImage);
                            logger.info("Added image to PDF for sheet: {}", sheet.getSheetName());
                        } catch (IOException e) {
                            logger.warn("Skipping invalid image in sheet {}: {}", sheet.getSheetName(), e.getMessage());
                        }
                    }
                }
            }

            document.close();
            ByteArrayResource resource = new ByteArrayResource(baos.toByteArray());
            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + getOutputFileName(file.getOriginalFilename()));
            headers.add(HttpHeaders.CACHE_CONTROL, "no-cache, no-store, must-revalidate");
            headers.add(HttpHeaders.PRAGMA, "no-cache");
            headers.add(HttpHeaders.EXPIRES, "0");

            return ResponseEntity.ok()
                    .headers(headers)
                    .contentLength(baos.size())
                    .contentType(MediaType.APPLICATION_PDF)
                    .body(resource);

        } catch (java.io.IOException e) {
            throw new RuntimeException("Error processing Excel to PDF conversion: " + e.getMessage(), e);
        }
    }

    private PageSize determinePageSize(Sheet sheet) {
        PrintSetup printSetup = sheet.getPrintSetup();
        boolean isLandscape = printSetup.getLandscape();
        PageSize pageSize = isLandscape ? PageSize.A4.rotate() : PageSize.A4;
        logger.debug("Sheet {} orientation: {}", sheet.getSheetName(), isLandscape ? "Landscape" : "Portrait");
        return pageSize;
    }

    private void createSheetTable(Document document, Sheet sheet, PdfFont font, PdfFont fontBold) {
        String sheetName = sheet.getSheetName();
        boolean isCalendar = false;

        int maxColumns = 1;
        Set<Integer> nonEmptyColumns = new HashSet<>();
        for (Row row : sheet) {
            if (row != null) {
                maxColumns = Math.max(maxColumns, row.getPhysicalNumberOfCells());
                for (int col = 0; col < row.getLastCellNum(); col++) {
                    org.apache.poi.ss.usermodel.Cell cell = row.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String cellValue = getCellValue(cell);
                    if (cellValue != null && !cellValue.replaceAll("[\\s\\u00A0\\u200B\\uFEFF]+", "").isEmpty()) {
                        nonEmptyColumns.add(col);
                    }
                }
            }
        }

        if (maxColumns >= 7) {
            Row headerRow = sheet.getRow(0);
            if (headerRow != null) {
                int dayCount = 0;
                for (int col = 0; col < 7; col++) {
                    String cellValue = getCellValue(headerRow.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK));
                    if (DAY_PATTERN.matcher(cellValue).matches()) {
                        dayCount++;
                        nonEmptyColumns.add(col);
                    }
                }
                if (dayCount >= 5) {
                    isCalendar = true;
                    nonEmptyColumns.clear();
                    for (int col = 0; col < 7; col++) {
                        nonEmptyColumns.add(col);
                    }
                }
            }
        }

        if (nonEmptyColumns.isEmpty()) {
            logger.debug("Skipping empty sheet: {}", sheetName);
            return;
        }

        List<Integer> columnIndices = new ArrayList<>(nonEmptyColumns);
        columnIndices.sort(Integer::compareTo);

        if (isCalendar) {
            String monthYearHeader = sheetName;
            String monthName = sheetName;
            Matcher matcher = MONTH_PATTERN.matcher(sheetName);
            if (matcher.find()) {
                monthName = matcher.group();
                switch (monthName.toLowerCase()) {
                    case "jan": monthName = "January"; break;
                    case "feb": monthName = "February"; break;
                    case "mar": monthName = "March"; break;
                    case "apr": monthName = "April"; break;
                    case "may": monthName = "May"; break;
                    case "jun": monthName = "June"; break;
                    case "jul": monthName = "July"; break;
                    case "aug": monthName = "August"; break;
                    case "sep": monthName = "September"; break;
                    case "oct": monthName = "October"; break;
                    case "nov": monthName = "November"; break;
                    case "dec": monthName = "December"; break;
                }
                monthYearHeader = monthName.toUpperCase() + " 2023";
            }
            document.add(new Paragraph(monthYearHeader)
                    .setFont(fontBold)
                    .setFontSize(16)
                    .setTextAlignment(TextAlignment.CENTER)
                    .setMarginBottom(10f));

            float[] columnWidths = new float[MAX_COLUMNS];
            float totalWidth = 0;
            for (int i = 0; i < MAX_COLUMNS; i++) {
                int col = columnIndices.get(i);
                int excelWidth = sheet.getColumnWidth(col);
                float pdfWidth = excelWidth / 256f * 7f;
                columnWidths[i] = pdfWidth;
                totalWidth += pdfWidth;
            }

            float maxTableWidth = document.getPdfDocument().getDefaultPageSize().getWidth() - 2 * MARGIN;
            float[] normalizedWidths = new float[MAX_COLUMNS];
            if (totalWidth > 0) {
                float scale = maxTableWidth / totalWidth;
                for (int i = 0; i < MAX_COLUMNS; i++) {
                    normalizedWidths[i] = columnWidths[i] * scale;
                }
            } else {
                for (int i = 0; i < MAX_COLUMNS; i++) {
                    normalizedWidths[i] = maxTableWidth / MAX_COLUMNS;
                }
            }

            Table table = new Table(normalizedWidths).useAllAvailableWidth();

            for (int i = 0; i < MAX_COLUMNS; i++) {
                table.addCell(new Cell()
                        .add(new Paragraph(DAY_HEADERS[i])
                                .setFont(fontBold)
                                .setFontSize(10))
                        .setTextAlignment(TextAlignment.CENTER)
                        .setBackgroundColor(HEADER_BG_COLOR)
                        .setBorder(new com.itextpdf.layout.borders.SolidBorder(1f))
                        .setPadding(5f));
            }

            List<String> allCellValues = new ArrayList<>();
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    for (int i = 0; i < MAX_COLUMNS; i++) {
                        int col = columnIndices.get(i);
                        org.apache.poi.ss.usermodel.Cell excelCell = row.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        String cellValue = getCellValue(excelCell);
                        if (cellValue != null && !cellValue.replaceAll("[\\s\\u00A0\\u200B\\uFEFF]+", "").isEmpty()) {
                            allCellValues.add(cellValue);
                        }
                    }
                }
            }

            int firstDayOfMonthIndex = -1;
            for (int i = 0; i < allCellValues.size(); i++) {
                String cellValue = allCellValues.get(i);
                String[] parts = cellValue.split("\n");
                String date = parts.length > 0 ? parts[0].trim() : "";
                try {
                    double serialDate = Double.parseDouble(date);
                    if (serialDate > 40000 && serialDate < 50000) {
                        int dayOffset = (int) (serialDate - 44926);
                        if (dayOffset == 1) {
                            firstDayOfMonthIndex = i % DAYS_IN_WEEK;
                            break;
                        }
                    }
                } catch (NumberFormatException e) {
                    continue;
                }
            }

            String[][] calendarGrid = new String[MAX_WEEKS][DAYS_IN_WEEK];
            int currentDay = 1;
            for (String cellValue : allCellValues) {
                String[] parts = cellValue.split("\n");
                String date = parts.length > 0 ? parts[0].trim() : "";
                String event = parts.length > 1 ? parts[1].trim() : "";
                try {
                    double serialDate = Double.parseDouble(date);
                    if (serialDate > 40000 && serialDate < 50000) {
                        int dayOffset = (int) (serialDate - 44926);
                        if (dayOffset == currentDay) {
                            int row = (firstDayOfMonthIndex + (dayOffset - 1)) / DAYS_IN_WEEK;
                            int col = (firstDayOfMonthIndex + (dayOffset - 1)) % DAYS_IN_WEEK;
                            if (row < MAX_WEEKS && col < DAYS_IN_WEEK) {
                                calendarGrid[row][col] = dayOffset + (event.isEmpty() ? "" : " " + event);
                            }
                            currentDay++;
                        }
                    }
                } catch (NumberFormatException e) {
                    continue;
                }
            }

            for (int row = 0; row < MAX_WEEKS; row++) {
                for (int col = 0; col < DAYS_IN_WEEK; col++) {
                    String cellContent = calendarGrid[row][col];
                    Cell cell = new Cell();
                    if (cellContent != null && !cellContent.isEmpty()) {
                        cell.add(new Paragraph(cellContent)
                                        .setFont(font)
                                        .setFontSize(10)
                                        .setTextAlignment(TextAlignment.CENTER))
                                .setTextAlignment(TextAlignment.CENTER)
                                .setBackgroundColor(row % 2 == 0 ? ROW_EVEN_BG_COLOR : ROW_ODD_BG_COLOR)
                                .setBorder(new com.itextpdf.layout.borders.SolidBorder(1f))
                                .setPadding(5f);
                    } else {
                        cell.setBackgroundColor(row % 2 == 0 ? ROW_EVEN_BG_COLOR : ROW_ODD_BG_COLOR)
                                .setBorder(Border.NO_BORDER)
                                .setPadding(5f);
                    }
                    table.addCell(cell);
                }
            }

            document.add(table.setMarginBottom(10f));
            logger.debug("Added calendar table for sheet {}: {} columns", sheetName, MAX_COLUMNS);

            for (String month : MONTHS_WITH_NOTES) {
                if (monthName.equalsIgnoreCase(month)) {
                    document.add(new Paragraph("Notes:")
                            .setFont(font)
                            .setFontSize(10)
                            .setTextAlignment(TextAlignment.RIGHT)
                            .setMarginTop(5f));
                    logger.debug("Added 'Notes:' section for sheet {}", sheetName);
                    break;
                }
            }
        } else {
            // Split columns into groups for normal tables
            List<List<Integer>> columnGroups = new ArrayList<>();
            for (int i = 0; i < columnIndices.size(); i += MAX_COLUMNS_PER_TABLE) {
                int end = Math.min(i + MAX_COLUMNS_PER_TABLE, columnIndices.size());
                columnGroups.add(new ArrayList<>(columnIndices.subList(i, end)));
            }

            for (int groupIndex = 0; groupIndex < columnGroups.size(); groupIndex++) {
                List<Integer> group = columnGroups.get(groupIndex);
                int numCols = group.size();

                float[] columnWidths = new float[numCols];
                float totalWidth = 0;
                for (int i = 0; i < numCols; i++) {
                    int col = group.get(i);
                    int excelWidth = sheet.getColumnWidth(col);
                    float pdfWidth = excelWidth / 256f * 7f;
                    columnWidths[i] = pdfWidth;
                    totalWidth += pdfWidth;
                }

                float maxTableWidth = document.getPdfDocument().getDefaultPageSize().getWidth() - 2 * MARGIN;
                float[] normalizedWidths = new float[numCols];
                if (totalWidth > 0) {
                    float scale = maxTableWidth / totalWidth;
                    for (int i = 0; i < numCols; i++) {
                        normalizedWidths[i] = columnWidths[i] * scale;
                    }
                } else {
                    for (int i = 0; i < numCols; i++) {
                        normalizedWidths[i] = maxTableWidth / numCols;
                    }
                }

                Table table = new Table(normalizedWidths).useAllAvailableWidth();

                Row headerRow = sheet.getRow(0);
                if (headerRow != null) {
                    for (int i = 0; i < numCols; i++) {
                        int col = group.get(i);
                        org.apache.poi.ss.usermodel.Cell excelCell = headerRow.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        String cellValue = getCellValue(excelCell);
                        Cell cell = new Cell()
                                .add(new Paragraph(cellValue)
                                        .setFont(fontBold)
                                        .setFontSize(10))
                                .setTextAlignment(TextAlignment.CENTER)
                                .setBackgroundColor(HEADER_BG_COLOR)
                                .setBorder(new com.itextpdf.layout.borders.SolidBorder(1f))
                                .setPadding(5f);
                        applyCellFormatting(cell, excelCell, font, fontBold);
                        table.addCell(cell);
                    }
                }

                int dataRowIndex = 0;
                for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        boolean isEmptyRow = true;
                        String[] cellValues = new String[numCols];
                        for (int i = 0; i < numCols; i++) {
                            int col = group.get(i);
                            org.apache.poi.ss.usermodel.Cell excelCell = row.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            cellValues[i] = getCellValue(excelCell);
                            if (cellValues[i] != null && !cellValues[i].replaceAll("[\\s\\u00A0\\u200B\\uFEFF]+", "").isEmpty()) {
                                isEmptyRow = false;
                            }
                        }
                        if (isEmptyRow) {
                            logger.debug("Skipping empty row {} in sheet {}", rowIndex, sheetName);
                            continue;
                        }
                        for (int j = 0; j < numCols; j++) {
                            int col = group.get(j);
                            org.apache.poi.ss.usermodel.Cell excelCell = row.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            String cellValue = cellValues[j];
                            Cell cell = new Cell()
                                    .add(new Paragraph(cellValue)
                                            .setFont(font)
                                            .setFontSize(10)
                                            .setTextAlignment(TextAlignment.CENTER))
                                    .setBackgroundColor(dataRowIndex % 2 == 0 ? ROW_EVEN_BG_COLOR : ROW_ODD_BG_COLOR)
                                    .setBorder(new com.itextpdf.layout.borders.SolidBorder(1f))
                                    .setPadding(5f);
                            applyCellFormatting(cell, excelCell, font, fontBold);
                            table.addCell(cell);
                        }
                        dataRowIndex++;
                    }
                }

                document.add(table.setMarginBottom(20f));
                logger.debug("Added table for sheet {}, group {}: {} columns", sheetName, groupIndex, numCols);
            }
        }
    }

    private void applyCellFormatting(Cell pdfCell, org.apache.poi.ss.usermodel.Cell excelCell, PdfFont font, PdfFont fontBold) {
        if (excelCell == null) return;

        CellStyle cellStyle = excelCell.getCellStyle();
        if (cellStyle != null) {
            Font fontStyle = excelCell.getSheet().getWorkbook().getFontAt(cellStyle.getFontIndexAsInt());
            if (fontStyle != null && fontStyle.getBold()) {
                for (Object element : pdfCell.getChildren()) {
                    if (element instanceof Paragraph) {
                        ((Paragraph) element).setFont(fontBold);
                    }
                }
            }

            org.apache.poi.ss.usermodel.Color color = cellStyle.getFillForegroundColorColor();
            if (color instanceof XSSFColor) {
                byte[] rgb = ((XSSFColor) color).getRGB();
                if (rgb != null && rgb.length == 3) {
                    pdfCell.setBackgroundColor(new DeviceRgb(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
                }
            }

            org.apache.poi.ss.usermodel.HorizontalAlignment alignment = cellStyle.getAlignment();
            if (alignment == org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER) {
                pdfCell.setTextAlignment(TextAlignment.CENTER);
            } else if (alignment == org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT) {
                pdfCell.setTextAlignment(TextAlignment.RIGHT);
            } else {
                pdfCell.setTextAlignment(TextAlignment.LEFT);
            }
        }
    }

    private String getCellValue(org.apache.poi.ss.usermodel.Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue() != null ? cell.getStringCellValue().trim() : "";
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    try {
                        SimpleDateFormat sdf = new SimpleDateFormat("'$'M/d/yyyy'$'");
                        return cell.getDateCellValue() != null ? sdf.format(cell.getDateCellValue()) : "";
                    } catch (IllegalStateException e) {
                        return "";
                    }
                }
                double numericValue = cell.getNumericCellValue();
                if (Math.abs(numericValue - Math.round(numericValue)) < 0.0001) {
                    return String.valueOf((int) numericValue);
                }
                return String.valueOf(numericValue);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getCellFormula() != null ? cell.getCellFormula() : "";
                } catch (IllegalStateException e) {
                    return "";
                }
            case ERROR:
                return "";
            default:
                return "";
        }
    }

    private List<byte[]> extractImagesFromSheet(Sheet sheet) {
        List<byte[]> images = new ArrayList<>();
        XSSFDrawing drawing = (XSSFDrawing) sheet.getDrawingPatriarch();
        if (drawing != null) {
            for (XSSFShape shape : drawing.getShapes()) {
                if (shape instanceof XSSFPicture) {
                    XSSFPicture picture = (XSSFPicture) shape;
                    XSSFPictureData pictureData = picture.getPictureData();
                    if (pictureData != null) {
                        byte[] imageData = pictureData.getData();
                        if (imageData != null && imageData.length > 0) {
                            String format = detectImageFormat(imageData);
                            if (format != null) {
                                images.add(imageData);
                                logger.debug("Found valid image in sheet {}: format={}, size={} bytes", sheet.getSheetName(), format, imageData.length);
                            } else {
                                logger.warn("Skipping invalid or unsupported image in sheet {}", sheet.getSheetName());
                            }
                        }
                    }
                }
            }
        }
        return images;
    }

    private String detectImageFormat(byte[] imageData) {
        if (imageData == null || imageData.length < 4) return null;
        try {
            if (imageData[0] == (byte) 0x89 && imageData[1] == (byte) 0x50 && imageData[2] == (byte) 0x4E && imageData[3] == (byte) 0x47) return "PNG";
            if (imageData[0] == (byte) 0xFF && imageData[1] == (byte) 0xD8) return "JPEG";
            if (imageData[0] == (byte) 0x47 && imageData[1] == (byte) 0x49 && imageData[2] == (byte) 0x46 && imageData[3] == (byte) 0x38) return "GIF";
            if (imageData[0] == (byte) 0x42 && imageData[1] == (byte) 0x4D) return "BMP";
            return null;
        } catch (Exception e) {
            logger.warn("Error detecting image format: {}", e.getMessage());
            return null;
        }
    }

    private String getOutputFileName(String originalFileName) {
        if (originalFileName == null || originalFileName.isEmpty()) return "output.pdf";
        return originalFileName.replaceAll("\\.xlsx$", ".pdf");
    }
}