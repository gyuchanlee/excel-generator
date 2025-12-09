package com.dodo.excelgenerator.excelgen.controller;

import com.dodo.excelgenerator.excelgen.dto.ExcelRequestDto;
import com.dodo.excelgenerator.excelgen.dto.ExcelResponseDto;
import com.dodo.excelgenerator.excelgen.service.ExcelService;
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

@Slf4j
@Controller
@RequestMapping("/excel")
@RequiredArgsConstructor
public class ExcelController {

    private final ExcelService excelService;
    private static final String SESSION_KEY = "excelData";

    /**
     * 메인 페이지
     */
    @GetMapping
    public String home(HttpSession session, Model model) {
        ExcelResponseDto data = (ExcelResponseDto) session.getAttribute(SESSION_KEY);
        if (data == null) {
            data = ExcelResponseDto.empty();
        }
        model.addAttribute("excelData", data);
        return "excelGen/home";
    }

    /**
     * 엑셀 파일 업로드
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