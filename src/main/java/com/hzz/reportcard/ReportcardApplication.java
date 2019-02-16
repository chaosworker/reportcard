package com.hzz.reportcard;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import java.io.*;
import java.util.*;


@SpringBootApplication
public class ReportcardApplication {

	public static void main(String[] args) {
		SpringApplication.run(ReportcardApplication.class, args);
	}

}

