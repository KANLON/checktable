package com.kanlon.tool;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/**
 * java操作的excel的工具类
 *
 * @author zhangcanlong
 * @date 2018年9月30日
 */
public class JExcelOption {
	/*
	 * public static void main(String[] args) { JExcelOption option = new
	 * JExcelOption(); // 本地excel地址 String path = "1.xls";
	 * System.out.println(option.readExcel(path));
	 *
	 * // List<HashMap<String, String>> list = new ArrayList<>(); // list =
	 * option.readExcel2(path); // for (int i = 0; i < list.size(); i++) { //
	 * System.out.println(list.get(i)); // } }
	 */

	/**
	 * 根据excel表位置，读取excel表格
	 *
	 * @param path
	 *            本地excel路径
	 * @return 返回嵌套list集合
	 */
	public List<ArrayList<String>> readExcel(String path) throws BiffException, IOException {
		File f = new File(path);
		Workbook book = Workbook.getWorkbook(f);//
		Sheet sheet = book.getSheet(0); // 获得第一个工作表对象
		// 存放所有excel表信息
		List<ArrayList<String>> list = new ArrayList<>();
		for (int i = 0; i < sheet.getRows(); i++) {
			// 存放一行的数据
			ArrayList<String> rowList = new ArrayList<>();
			for (int j = 0; j < sheet.getColumns(); j++) {
				Cell cell = sheet.getCell(j, i); // 获得单元格
				rowList.add(cell.getContents());
				// System.out.print(cell.getContents() + " ");
				// 得到单元格的类型
				// System.out.println(cell.getType());
			}
			list.add(rowList);
			// System.out.print("\n");
		}
		return list;
	}

	/**
	 * 根据excel表位置，读取excel表格
	 *
	 * @param path
	 *            本地excel路径
	 * @return 返回每行的map集合
	 * @throws IOException
	 * @throws BiffException
	 */
	public List<HashMap<String, String>> readExcel2(String path) throws BiffException, IOException {
		File f = new File(path);
		Workbook book = Workbook.getWorkbook(f);//
		Sheet sheet = book.getSheet(0); // 获得第一个工作表对象
		// 存放所有excel表信息
		List<HashMap<String, String>> list = new ArrayList<>();
		for (int i = 1; i < sheet.getRows(); i++) {
			// 存放一行的数据
			HashMap<String, String> map = new HashMap<>();
			for (int j = 0; j < sheet.getColumns(); j++) {
				Cell cell = sheet.getCell(j, i); // 获得单元格
				map.put(sheet.getCell(j, 0).getContents(), cell.getContents());
				// System.out.print(cell.getContents() + " ");
				// 得到单元格的类型
				// System.out.println(cell.getType());
			}
			list.add(map);
			// System.out.print("\n");
		}
		return list;
	}

	/**
	 * 将 嵌套list集合输出到excel中
	 *
	 * @param outList
	 *            要输出的list集合
	 * @param targetExcelPath
	 *            要输出到那个excel中
	 * @throws RowsExceededException
	 * @throws WriteException
	 * @throws IOException
	 */
	public boolean writeExcel(List<ArrayList<String>> outList, String targetExcelPath)
			throws RowsExceededException, WriteException, IOException {
		if (outList == null || outList.size() == 0) {
			return true;
		}

		WritableWorkbook workbook = null;
		workbook = Workbook.createWorkbook(new File(targetExcelPath));
		// 生成第一页的工作表，参数为0说明是第一页
		WritableSheet sheet = workbook.createSheet("纸质版识别内容及核对结果", 0);
		for (int i = 0; i < outList.size(); i++) {
			for (int j = 0; j < outList.get(i).size(); j++) {
				ArrayList<String> tempList = outList.get(i);
				Label label = new Label(j, i, tempList.get(j));
				sheet.addCell(label);
			}
		}
		// 数字类型
		// jxl.write.Number number = new jxl.write.Number(0,1,789.123);
		workbook.write();
		workbook.close();
		return true;
	}

}