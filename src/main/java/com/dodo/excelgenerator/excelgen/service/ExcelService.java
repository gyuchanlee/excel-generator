package com.dodo.excelgenerator.excelgen.service;

import com.dodo.excelgenerator.excelgen.dto.ExcelResponseDto;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Slf4j
@Service
public class ExcelService {

    /**
     * 엑셀 파일 파싱
     */
    public ExcelResponseDto parseExcel(MultipartFile file) throws IOException {
        List<String> headers = new ArrayList<>();
        List<List<String>> rows = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);
            boolean isFirstRow = true;

            for (Row row : sheet) {
                List<String> rowData = new ArrayList<>();
                for (Cell cell : row) {
                    rowData.add(getCellValueAsString(cell));
                }

                if (isFirstRow) {
                    headers = rowData;
                    isFirstRow = false;
                } else {
                    // 빈 행이 아닌 경우만 추가
                    if (rowData.stream().anyMatch(s -> s != null && !s.trim().isEmpty())) {
                        rows.add(rowData);
                    }
                }
            }
        }

        return ExcelResponseDto.builder()
                .headers(headers)
                .rows(rows)
                .totalRows(rows.size())
                .fileName(file.getOriginalFilename())
                .build();
    }

    /**
     * 컬럼(헤더) 일치 여부 검증
     */
    public boolean validateHeaders(List<String> baseHeaders, List<String> newHeaders) {
        if (baseHeaders.size() != newHeaders.size()) {
            return false;
        }
        for (int i = 0; i < baseHeaders.size(); i++) {
            if (!baseHeaders.get(i).equals(newHeaders.get(i))) {
                return false;
            }
        }
        return true;
    }

    /**
     * 기존 데이터에 새 데이터 병합 (헤더 제외, 데이터 행만 추가)
     */
    public ExcelResponseDto mergeData(ExcelResponseDto base, ExcelResponseDto newData) {
        List<List<String>> mergedRows = new ArrayList<>(base.getRows());
        mergedRows.addAll(newData.getRows());

        return ExcelResponseDto.builder()
                .headers(base.getHeaders())
                .rows(mergedRows)
                .totalRows(mergedRows.size())
                .fileName(base.getFileName())
                .build();
    }

    /**
     * 데이터를 엑셀 파일로 생성
     */
    public byte[] createExcel(ExcelResponseDto data) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.createSheet("Data");

            // 헤더 스타일
            CellStyle headerStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // 헤더 행 생성
            Row headerRow = sheet.createRow(0);
            List<String> headers = data.getHeaders();
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers.get(i));
                cell.setCellStyle(headerStyle);
            }

            // 데이터 행 생성
            List<List<String>> rows = data.getRows();
            for (int i = 0; i < rows.size(); i++) {
                Row row = sheet.createRow(i + 1);
                List<String> rowData = rows.get(i);
                for (int j = 0; j < rowData.size(); j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(rowData.get(j));
                }
            }

            // 컬럼 너비 자동 조정 + 최소 너비 설정 > 최소, 최대 넓이 확인하면서 조절
            int minWidth = 15 * 256;
            int maxWidth = 50 * 256;

            for (int i = 0; i < headers.size(); i++) {
                sheet.autoSizeColumn(i);
                int currentWidth = sheet.getColumnWidth(i);

                if (currentWidth < minWidth) {
                    sheet.setColumnWidth(i, minWidth);
                } else if (currentWidth > maxWidth) {
                    sheet.setColumnWidth(i, maxWidth);
                }
            }

            workbook.write(out);
            return out.toByteArray();
        }
    }

    /**
     * 셀 값을 문자열로 변환
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    yield cell.getLocalDateTimeCellValue().toString();
                }
                // 정수면 소수점 없이 출력
                double value = cell.getNumericCellValue();
                if (value == Math.floor(value)) {
                    yield String.valueOf((long) value);
                }
                yield String.valueOf(value);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> {
                try {
                    yield String.valueOf(cell.getNumericCellValue());
                } catch (Exception e) {
                    yield cell.getStringCellValue();
                }
            }
            case BLANK -> "";
            default -> "";
        };
    }
}