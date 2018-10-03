package com.kanlon.test;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Properties;

import com.kanlon.tool.CustomExceptionTool;
import com.kanlon.tool.JExcelOption;

public class Test {

	public static void main(String[] args) {
		OcrClient ocrClient = new OcrClient();
		JExcelOption option = new JExcelOption();

		Properties prop = new Properties();
		try {
			prop.load(new InputStreamReader(new FileInputStream(".\\info.properties"), "UTF-8"));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		String paperPath = prop.getProperty("target_path");
		String elecPath = prop.getProperty("excel_path");
		try {
			System.out.println(option.compareAwardTable(paperPath, elecPath));
		} catch (Exception e) {
			throw new RuntimeException(
					"纸质版信息表格时和电子版信息表格是发送错误！请确认关闭所有相应excel表后重试！\n" + CustomExceptionTool.getExceptionMsg(e));
		}
	}

}
