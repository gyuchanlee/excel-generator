package com.dodo.excelgenerator.excelgen.controller;

import com.dodo.excelgenerator.excelgen.dto.ExcelRequestDto;
import com.dodo.excelgenerator.excelgen.dto.ExcelResponseDto;
import com.dodo.excelgenerator.excelgen.dto.TemplateConfigDto;
import com.dodo.excelgenerator.excelgen.service.ExcelService;
import com.dodo.excelgenerator.excelgen.service.TemplateParsingService;
import jakarta.servlet.http.HttpSession;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

@Slf4j
@Controller
@RequestMapping("/excel")
@RequiredArgsConstructor
public class ExcelController {

    private final ExcelService excelService;
    private final TemplateParsingService templateParsingService;

    private static final String SESSION_KEY = "excelData";
    private static final String CONFIG_KEY = "templateConfig";

    /**
     * 메인 페이지
     */
    @GetMapping
    public String home(HttpSession session, Model model) {
        ExcelResponseDto data = (ExcelResponseDto) session.getAttribute(SESSION_KEY);
        if (data == null) {
            data = ExcelResponseDto.empty();
        }

        // 템플릿 설정 (없으면 기본값)
        TemplateConfigDto config = (TemplateConfigDto) session.getAttribute(CONFIG_KEY);
        if (config == null) {
            config = TemplateConfigDto.defaultConfig();
        }

        model.addAttribute("excelData", data);
        model.addAttribute("config", config);
        return "excelGen/home";
    }

    // ===================================================================
    // 일반 엑셀 업로드 (동일 구조 엑셀 병합)
    // ===================================================================

    /**
     * 일반 엑셀 - 다중 파일 업로드 및 병합
     */
    @PostMapping("/upload-multiple")
    public String uploadMultiple(@RequestParam("files") List<MultipartFile> files,
                                 HttpSession session,
                                 RedirectAttributes redirectAttributes) {

        // 빈 파일 체크
        if (files == null || files.isEmpty() || files.stream().allMatch(MultipartFile::isEmpty)) {
            redirectAttributes.addFlashAttribute("error", "파일을 선택해주세요.");
            return "redirect:/excel";
        }

        // 유효한 파일만 필터링
        List<MultipartFile> validFiles = files.stream()
                .filter(f -> !f.isEmpty())
                .toList();

        if (validFiles.isEmpty()) {
            redirectAttributes.addFlashAttribute("error", "유효한 파일이 없습니다.");
            return "redirect:/excel";
        }

        int successCount = 0;
        int failCount = 0;
        List<String> failedFiles = new ArrayList<>();
        ExcelResponseDto mergedData = (ExcelResponseDto) session.getAttribute(SESSION_KEY);

        for (MultipartFile file : validFiles) {
            try {
                ExcelResponseDto newData = excelService.parseExcel(file);

                if (mergedData == null || mergedData.getHeaders().isEmpty()) {
                    // 첫 번째 파일 - 기준 데이터로 설정
                    mergedData = newData;
                    successCount++;
                } else {
                    // 이후 파일들 - 헤더 검증 후 병합
                    if (excelService.validateHeaders(mergedData.getHeaders(), newData.getHeaders())) {
                        mergedData = excelService.mergeData(mergedData, newData);
                        successCount++;
                    } else {
                        failCount++;
                        failedFiles.add(file.getOriginalFilename() + " (컬럼 불일치)");
                    }
                }
            } catch (IOException e) {
                log.error("파일 처리 실패: {}", file.getOriginalFilename(), e);
                failCount++;
                failedFiles.add(file.getOriginalFilename() + " (처리 오류)");
            }
        }

        // 세션에 저장
        if (mergedData != null && !mergedData.getHeaders().isEmpty()) {
            session.setAttribute(SESSION_KEY, mergedData);
        }

        // 결과 메시지 생성
        StringBuilder message = new StringBuilder();
        message.append(String.format("✅ %d개 파일 병합 완료 (총 %d행)",
                successCount, mergedData != null ? mergedData.getTotalRows() : 0));

        if (failCount > 0) {
            message.append(String.format("\n⚠️ %d개 파일 실패", failCount));
            redirectAttributes.addFlashAttribute("failedFiles", failedFiles);
        }

        redirectAttributes.addFlashAttribute("message", message.toString());

        return "redirect:/excel";
    }

    /**
     * 일반 엑셀 - 단일 파일 업로드
     */
    @PostMapping("/upload")
    public String upload(@RequestParam("file") MultipartFile file,
                         HttpSession session,
                         RedirectAttributes redirectAttributes) {

        if (file.isEmpty()) {
            redirectAttributes.addFlashAttribute("error", "파일을 선택해주세요.");
            return "redirect:/excel";
        }

        try {
            ExcelResponseDto newData = excelService.parseExcel(file);
            ExcelResponseDto existingData = (ExcelResponseDto) session.getAttribute(SESSION_KEY);

            if (existingData == null || existingData.getHeaders().isEmpty()) {
                // 첫 번째 파일 업로드
                session.setAttribute(SESSION_KEY, newData);
                redirectAttributes.addFlashAttribute("message",
                        "파일이 업로드되었습니다. (" + newData.getTotalRows() + "행)");
            } else {
                // 추가 파일 업로드 - 헤더 검증 후 병합
                if (excelService.validateHeaders(existingData.getHeaders(), newData.getHeaders())) {
                    ExcelResponseDto mergedData = excelService.mergeData(existingData, newData);
                    session.setAttribute(SESSION_KEY, mergedData);
                    redirectAttributes.addFlashAttribute("message",
                            "데이터가 병합되었습니다. (총 " + mergedData.getTotalRows() + "행)");
                } else {
                    redirectAttributes.addFlashAttribute("error",
                            "컬럼 구조가 다릅니다. 같은 형식의 파일을 업로드해주세요.");
                }
            }
        } catch (IOException e) {
            log.error("파일 업로드 실패", e);
            redirectAttributes.addFlashAttribute("error", "파일 처리 중 오류가 발생했습니다.");
        }

        return "redirect:/excel";
    }

    // ===================================================================
    // 템플릿 파싱 업로드 (피벗 + 분리 테이블 → 통합)
    // ===================================================================

    /**
     * 템플릿 설정 저장
     */
    @PostMapping("/config")
    public String saveConfig(@RequestParam("companyRow") int companyRow,
                             @RequestParam("companyCol") int companyCol,
                             @RequestParam("codeRow") int codeRow,
                             @RequestParam("codeCol") int codeCol,
                             @RequestParam("dataStartRow") int dataStartRow,
                             @RequestParam("leftTableStartCol") int leftTableStartCol,
                             @RequestParam("rightTableStartCol") int rightTableStartCol,
                             @RequestParam("colCount") int colCount,
                             HttpSession session,
                             RedirectAttributes redirectAttributes) {

        TemplateConfigDto config = TemplateConfigDto.builder()
                .companyRow(companyRow)
                .companyCol(companyCol)
                .codeRow(codeRow)
                .codeCol(codeCol)
                .dataStartRow(dataStartRow)
                .leftTableStartCol(leftTableStartCol)
                .rightTableStartCol(rightTableStartCol)
                .colCount(colCount)
                .build();

        session.setAttribute(CONFIG_KEY, config);
        redirectAttributes.addFlashAttribute("message", "✅ 템플릿 설정이 저장되었습니다.");

        return "redirect:/excel";
    }

    /**
     * 템플릿 다중 파일 업로드 (커스텀 파싱)
     */
    @PostMapping("/upload-template")
    public String uploadTemplate(@RequestParam("files") List<MultipartFile> files,
                                 HttpSession session,
                                 RedirectAttributes redirectAttributes) {

        // 빈 파일 체크
        if (files == null || files.isEmpty() || files.stream().allMatch(MultipartFile::isEmpty)) {
            redirectAttributes.addFlashAttribute("error", "파일을 선택해주세요.");
            return "redirect:/excel";
        }

        // 유효한 파일만 필터링
        List<MultipartFile> validFiles = files.stream()
                .filter(f -> !f.isEmpty())
                .toList();

        // 템플릿 설정 가져오기
        TemplateConfigDto config = (TemplateConfigDto) session.getAttribute(CONFIG_KEY);
        if (config == null) {
            config = TemplateConfigDto.defaultConfig();
        }

        int successCount = 0;
        int failCount = 0;
        List<String> failedFiles = new ArrayList<>();
        ExcelResponseDto mergedData = (ExcelResponseDto) session.getAttribute(SESSION_KEY);

        for (MultipartFile file : validFiles) {
            try {
                ExcelResponseDto newData = templateParsingService.parseTemplate(file, config);

                if (mergedData == null || mergedData.getHeaders().isEmpty()) {
                    // 첫 번째 파일
                    mergedData = newData;
                } else {
                    // 이후 파일들 병합 (헤더는 동일하므로 바로 병합)
                    mergedData = excelService.mergeData(mergedData, newData);
                }
                successCount++;

            } catch (IOException e) {
                log.error("파일 처리 실패: {}", file.getOriginalFilename(), e);
                failCount++;
                failedFiles.add(file.getOriginalFilename() + " (처리 오류)");
            }
        }

        // 세션에 저장
        if (mergedData != null && !mergedData.getHeaders().isEmpty()) {
            session.setAttribute(SESSION_KEY, mergedData);
        }

        // 결과 메시지
        StringBuilder message = new StringBuilder();
        message.append(String.format("✅ %d개 파일 파싱 완료 (총 %d행)",
                successCount, mergedData != null ? mergedData.getTotalRows() : 0));

        if (failCount > 0) {
            message.append(String.format("\n⚠️ %d개 파일 실패", failCount));
            redirectAttributes.addFlashAttribute("failedFiles", failedFiles);
        }

        redirectAttributes.addFlashAttribute("message", message.toString());

        return "redirect:/excel";
    }

    /**
     * 설정 초기화
     */
    @PostMapping("/config/reset")
    public String resetConfig(HttpSession session, RedirectAttributes redirectAttributes) {
        session.setAttribute(CONFIG_KEY, TemplateConfigDto.defaultConfig());
        redirectAttributes.addFlashAttribute("message", "설정이 기본값으로 초기화되었습니다.");
        return "redirect:/excel";
    }

    // ===================================================================
    // 공통 기능 (수정, 다운로드, 초기화)
    // ===================================================================

    /**
     * 테이블 데이터 수정 (AJAX)
     */
    @PostMapping("/update")
    @ResponseBody
    public ResponseEntity<String> update(@RequestBody ExcelRequestDto request,
                                         HttpSession session) {
        ExcelResponseDto data = ExcelResponseDto.builder()
                .headers(request.getHeaders())
                .rows(request.getRows())
                .totalRows(request.getRows().size())
                .fileName("merged_data")
                .build();

        session.setAttribute(SESSION_KEY, data);
        return ResponseEntity.ok("저장되었습니다.");
    }

    /**
     * 엑셀 파일 다운로드
     */
    @GetMapping("/download")
    public ResponseEntity<byte[]> download(HttpSession session) throws IOException {
        ExcelResponseDto data = (ExcelResponseDto) session.getAttribute(SESSION_KEY);

        if (data == null || data.getRows().isEmpty()) {
            return ResponseEntity.badRequest().build();
        }

        byte[] excelBytes = excelService.createExcel(data);
        String fileName = URLEncoder.encode("merged_excel_data.xlsx", StandardCharsets.UTF_8);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + fileName + "\"")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .contentLength(excelBytes.length)
                .body(excelBytes);
    }

    /**
     * 데이터 초기화
     */
    @PostMapping("/clear")
    public String clear(HttpSession session, RedirectAttributes redirectAttributes) {
        session.removeAttribute(SESSION_KEY);
        redirectAttributes.addFlashAttribute("message", "데이터가 초기화되었습니다.");
        return "redirect:/excel";
    }
}