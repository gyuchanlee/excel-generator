package com.dodo.excelgenerator.excelgen.service;

import com.dodo.excelgenerator.excelgen.dto.ExcelResponseDto;
import com.dodo.excelgenerator.excelgen.dto.TemplateConfigDto;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * 커스텀 템플릿 엑셀 파싱 서비스
 * - 피벗 테이블에서 회사/코드 추출
 * - 왼쪽 테이블 + 오른쪽 테이블을 가로로 합쳐서 한 행으로 만듦
 */
@Slf4j
@Service
public class TemplateParsingService {

    /**
     * 템플릿 엑셀 파싱 (설정 기반)
     */
    public ExcelResponseDto parseTemplate(MultipartFile file, TemplateConfigDto config) throws IOException {
        List<List<String>> allRows = new ArrayList<>();
        List<String> headers = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);

            // 1. 피벗 테이블에서 회사/코드 추출
            String company = getCellValue(sheet, config.getCompanyRow(), config.getCompanyCol());
            String code = getCellValue(sheet, config.getCodeRow(), config.getCodeCol());

            log.info("추출된 피벗 데이터 - 회사: {}, 코드: {}", company, code);

            // 2. 헤더 구성: 코드, 회사, [왼쪽 테이블 헤더들], [오른쪽 테이블 헤더들]
            headers.add("코드");
            headers.add("회사");

            // 왼쪽 테이블 헤더 추출
            List<String> leftHeaders = getRowData(sheet, config.getDataStartRow(), config.getLeftTableStartCol(), config.getColCount());
            headers.addAll(leftHeaders);

            // 오른쪽 테이블 헤더 추출
            List<String> rightHeaders = getRowData(sheet, config.getDataStartRow(), config.getRightTableStartCol(), config.getColCount());
            headers.addAll(rightHeaders);

            log.info("헤더 구성: {}", headers);

            // 3. 데이터 행 파싱 (헤더 다음 행부터)
            int dataRowStart = config.getDataStartRow() + 1;
            int currentRow = dataRowStart;

            while (currentRow <= sheet.getLastRowNum()) {
                Row row = sheet.getRow(currentRow);
                if (row == null) {
                    currentRow++;
                    continue;
                }

                // 왼쪽 테이블 데이터
                List<String> leftData = getRowData(sheet, currentRow, config.getLeftTableStartCol(), config.getColCount());

                // 왼쪽 테이블이 비어있으면 종료
                if (isEmptyRow(leftData)) {
                    break;
                }

                // 오른쪽 테이블 데이터
                List<String> rightData = getRowData(sheet, currentRow, config.getRightTableStartCol(), config.getColCount());

                // 한 행으로 합치기: 코드 + 회사 + 왼쪽 데이터 + 오른쪽 데이터
                List<String> mergedRow = new ArrayList<>();
                mergedRow.add(code);
                mergedRow.add(company);
                mergedRow.addAll(leftData);
                mergedRow.addAll(rightData);

                allRows.add(mergedRow);
                log.debug("행 {}: {}", currentRow, mergedRow);

                currentRow++;
            }

            log.info("총 파싱된 행 수: {}", allRows.size());
        }

        return ExcelResponseDto.builder()
                .headers(headers)
                .rows(allRows)
                .totalRows(allRows.size())
                .fileName(file.getOriginalFilename())
                .build();
    }

    /**
     * 특정 행의 데이터를 지정된 열부터 colCount만큼 가져오기
     */
    private List<String> getRowData(Sheet sheet, int rowIdx, int startCol, int colCount) {
        List<String> data = new ArrayList<>();
        Row row = sheet.getRow(rowIdx);

        if (row == null) {
            // 빈 행이면 빈 문자열로 채움
            for (int i = 0; i < colCount; i++) {
                data.add("");
            }
            return data;
        }

        for (int i = 0; i < colCount; i++) {
            Cell cell = row.getCell(startCol + i);
            data.add(getCellValueAsString(cell));
        }

        return data;
    }

    /**
     * 행이 비어있는지 확인
     */
    private boolean isEmptyRow(List<String> rowData) {
        return rowData.stream().allMatch(s -> s == null || s.trim().isEmpty());
    }

    /**
     * 특정 셀 값 가져오기
     */
    private String getCellValue(Sheet sheet, int rowIdx, int colIdx) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) return "";

        Cell cell = row.getCell(colIdx);
        return getCellValueAsString(cell);
    }

    /**
     * 셀 값을 문자열로 변환
     */
    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    yield cell.getLocalDateTimeCellValue().toString();
                }
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