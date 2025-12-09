package com.dodo.excelgenerator.excelgen.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.ArrayList;
import java.util.List;

/**
 * 엑셀 데이터 담기용 Dto - 헤더와 행 데이터 저장
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelResponseDto {

    private List<String> headers;           // 첫 번째 행 (컬럼명)
    private List<List<String>> rows;        // 데이터 행들
    private int totalRows;                  // 전체 행 수
    private String fileName;                // 원본 파일명

    public static ExcelResponseDto empty() {
        return ExcelResponseDto.builder()
                .headers(new ArrayList<>())
                .rows(new ArrayList<>())
                .totalRows(0)
                .fileName("")
                .build();
    }
}