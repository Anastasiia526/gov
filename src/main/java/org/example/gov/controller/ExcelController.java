package org.example.gov.controller;

import org.example.gov.entity.WorkbookJson;
import org.example.gov.service.ExcelToJson;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import tools.jackson.databind.ObjectMapper;

import java.io.ByteArrayInputStream;

@Controller
public class ExcelController {

    private final ExcelToJson service;
    private final ObjectMapper objectMapper;

    public ExcelController(ExcelToJson service, ObjectMapper objectMapper) {
        this.service = service;
        this.objectMapper = objectMapper;
    }

    @GetMapping("/")
    public String index() {
        return "index";
    }

    @PostMapping("/preview")
    public String preview(@RequestParam("file") MultipartFile file, Model model) {
        WorkbookJson workbook = service.convert(file);
        model.addAttribute("workbook", workbook);
        return "preview";
    }

    @PostMapping("/convert")
    public ResponseEntity<InputStreamResource> convert(@RequestParam("file") MultipartFile file)
            throws RuntimeException {

        WorkbookJson workbook = service.convert(file);

        byte[] jsonBytes = objectMapper.writerWithDefaultPrettyPrinter()
                .writeValueAsBytes(workbook);

        InputStreamResource resource = new InputStreamResource(new ByteArrayInputStream(jsonBytes));

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION,
                        ContentDisposition.attachment()
                                .filename("result.json")
                                .build().toString())
                .contentType(MediaType.APPLICATION_JSON)
                .contentLength(jsonBytes.length)
                .body(resource);
    }

    @ExceptionHandler(Exception.class)
    public String handleError(Exception ex, Model model) {
        model.addAttribute("error", ex.getMessage());
        return "error";
    }
}
