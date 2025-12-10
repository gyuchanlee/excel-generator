package com.dodo.excelgenerator.excelgen.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 엑셀 템플릿 파싱 설정
 * - 피벗 테이블 위치 (회사, 코드)
 * - 데이터 테이블 시작 위치
 * - 오른쪽 테이블 시작 열
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class TemplateConfigDto {

    // 피벗 테이블 설정 (0-based index)
    private int companyRow;         // 회사 행 (기본: 0)
    private int companyCol;         // 회사 값 열 (기본: 1, B열)
    private int codeRow;            // 코드 행 (기본: 1)
    private int codeCol;            // 코드 값 열 (기본: 1, B열)

    // 데이터 테이블 설정
    private int dataStartRow;       // 데이터 테이블 시작 행 (기본: 3, 4행)
    private int leftTableStartCol;  // 왼쪽 테이블 시작 열 (기본: 0, A열)
    private int rightTableStartCol; // 오른쪽 테이블 시작 열 (기본: 4, E열)
    private int colCount;           // 각 테이블 컬럼 수 (기본: 3, 이름/부서/직급)

    /**
     * 기본 설정값
     */
    public static TemplateConfigDto defaultConfig() {
        return TemplateConfigDto.builder()
                .companyRow(0)
                .companyCol(1)
                .codeRow(1)
                .codeCol(1)
                .dataStartRow(3)
                .leftTableStartCol(0)
                .rightTableStartCol(4)
                .colCount(3)
                .build();
    }
}