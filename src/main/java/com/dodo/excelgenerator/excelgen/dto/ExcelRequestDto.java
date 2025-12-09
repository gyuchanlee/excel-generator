package com.dodo.excelgenerator.excelgen.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * 엑셀 수정 요청 - 헤더 및 데이터 수정
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExcelRequestDto {

    private List<String> headers;           // 수정된 헤더
    private List<List<String>> rows;        // 수정된 데이터 행들
}