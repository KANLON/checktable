package com.kanlon.test;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.TreeMap;
import java.util.logging.FileHandler;
import java.util.logging.Formatter;
import java.util.logging.Level;
import java.util.logging.LogRecord;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.imageio.stream.FileImageInputStream;

import org.apache.commons.lang.StringUtils;
import org.json.JSONArray;
import org.json.JSONObject;

import com.baidu.aip.ocr.AipOcr;
import com.kanlon.tool.CustomExceptionTool;
import com.kanlon.tool.JExcelOption;

import jxl.read.biff.BiffException;

/**
 * 百度API调用例子
 *
 * @author zhangcanlong
 * @date 2018年9月28日
 */
public class OcrClient {
	// 百度云应用注册信息(表格核对应用)
	public static final String APP_ID = "14310703";
	public static final String API_KEY = "LtY5IkZfQj3jmWk61hgc77N9";
	public static final String SECRET_KEY = "szug9FznvS8BHU0srS6IlieRUccGhRwd";
	// 各个字段的默认值
	public static final String TITLE_DEFAULT = "标题";
	public static final String SCHOOL_DEFAULT = "学校";
	public static final String DEPARTMENT_DEFAULT = "学院";
	public static final String MAJOR_DEFAULT = "专业";
	public static final String NAME_DEFAULT = "姓名";
	public static final String SEX_DEFAULT = "性别";
	public static final String NATION_DEFAULT = "民族";
	public static final String COMEDATE_DEFAULT = "入学年月";
	public static final String STUDENTID_DEFAULT = "学号";
	public static final String CLASSSTR_DEFAULT = "班级";
	public static final String PHONE_DEFAULT = "联系电话";
	public static final String RIGHTID_DEFAULT = "身份证号";
	public static final String POORLEVEL_DEFAULT = "家庭经济困难学生认定等级";
	public static final String TOTALNUM_DEFAULT = "总人数";
	public static final String STURANK_DEFAULT = "学习成绩排名";
	public static final String COMRANK_DEFAULT = "综合考评排名";
	public static final String SIGNDATE_DAFAULT = "个人签名日期";
	public static final String DEPARTMENTSIGNDATE_DEFAULT = "学院盖章日期";
	public static final String SCHOOLSIGNDATE_DEFAULT = "学校盖章日期";

	Properties prop = new Properties();
	private static Logger logger = Logger.getLogger(OcrClient.class.getName());
	/**
	 * 项目根目录，如果是打包成jar，则是cmd当前的目录
	 */
	public final static String projectPath = new File("./").getAbsolutePath();
	static {
		try {
			// 设置在控制台输出
			logger.setUseParentHandlers(true);
			// 设置日志输出等级
			logger.setLevel(Level.INFO);
			FileHandler fileHandler = null;
			fileHandler = new FileHandler(
					projectPath + "/logs/" + new SimpleDateFormat("yyyy_MM_dd_HH_mm_ss").format(new Date()) + ".log");
			final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
			fileHandler.setFormatter(new Formatter() {
				@Override
				public String format(LogRecord arg0) {
					return String.format("%-8s", arg0.getLevel().getLocalizedName())
							+ sdf.format(new Date(arg0.getMillis())) + "  : " + arg0.getMessage() + "\n";
				}
			});
			logger.addHandler(fileHandler);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		Properties prop = new Properties();
		try {
			// 打包成jar时的配置文件的位置，即cmd当前的位置,也可以直接在eclipse运行
			String jarPath = projectPath + "/info.properties";
			// prop.load(new FileReader(jarPath));
			prop.load(new InputStreamReader(new FileInputStream(jarPath), "UTF-8"));
		} catch (IOException e) {
			e.printStackTrace();
			logger.log(Level.SEVERE, "找不到文件名，加载配置失败！！" + CustomExceptionTool.getExceptionMsg(e));
			System.exit(0);
		}
		/**
		 * 识别的次数
		 */
		int num = 0;
		/**
		 * 记录错误的序号及错误原因的map集合
		 */
		Map<String, String> errorMap = new TreeMap<>();

		/**
		 * 记录与电子版核对的错误的序号及错误原因的map集合
		 */
		Map<String, String> checkErrorMap = new TreeMap<>();

		String appId = prop.getProperty("app_id", APP_ID);
		String apiKey = prop.getProperty("api_key", API_KEY);
		String secretKey = prop.getProperty("secret_key", SECRET_KEY);
		AipOcr client = new AipOcr(appId, apiKey, secretKey);
		OcrClient ocrClient = new OcrClient();
		// 要识别照片的存放目录或要识别的照片
		File fileRoot = new File(prop.getProperty("image_path", "压缩后的奖学金"));
		// 本地excel信息的存放地址
		String excelPath = prop.getProperty("excel_path", "要核对excel表信息.xls");
		JExcelOption option = new JExcelOption();
		List<HashMap<String, String>> list = new ArrayList<>();
		try {
			list = option.readExcel2(excelPath);
		} catch (BiffException | IOException e) {
			logger.log(Level.WARNING, "找不到电子版excel信息的存放地址\r\n" + CustomExceptionTool.getExceptionMsg(e));
		}

		// 要存放识别内容的excel
		String targetExcelPath = prop.getProperty("target_path", "paper_info.xls");
		// 存放识别内容的list集合
		List<ArrayList<String>> paperList = new ArrayList<>();
		// 存放扫描得到的学生信息(txt文件)
		File fileResult = new File("student_info.txt");

		File[] files = new File[1];
		if (fileRoot.isDirectory()) {
			files = fileRoot.listFiles();
		} else {
			files[0] = fileRoot;
		}

		FileWriter fw = null;
		try {
			fw = new FileWriter(fileResult);
		} catch (IOException e) {
			e.printStackTrace();
			logger.log(Level.SEVERE, "创建存放学生信息的txt错误！！" + CustomExceptionTool.getExceptionMsg(e));
		}

		StringBuffer titleBuffer = new StringBuffer("");
		titleBuffer.append("序号" + "\t" + "文件名" + "\t");
		titleBuffer.append(TITLE_DEFAULT + "\t");
		titleBuffer.append(SCHOOL_DEFAULT + "\t");
		titleBuffer.append(CLASSSTR_DEFAULT + "\t");
		titleBuffer.append(SIGNDATE_DAFAULT + "\t");
		titleBuffer.append(DEPARTMENTSIGNDATE_DEFAULT + "\t");
		titleBuffer.append(SCHOOLSIGNDATE_DEFAULT + "\t");
		titleBuffer.append(NAME_DEFAULT + "\t");
		titleBuffer.append(RIGHTID_DEFAULT + "\t");
		titleBuffer.append(DEPARTMENT_DEFAULT + "\t");
		titleBuffer.append(MAJOR_DEFAULT + "\t");
		titleBuffer.append(STUDENTID_DEFAULT + "\t");
		titleBuffer.append(SEX_DEFAULT + "\t");
		titleBuffer.append(NATION_DEFAULT + "\t");
		titleBuffer.append(COMEDATE_DEFAULT + "\t");
		titleBuffer.append(POORLEVEL_DEFAULT + "\t");
		titleBuffer.append(STURANK_DEFAULT + "\t");
		titleBuffer.append(TOTALNUM_DEFAULT + "\t");
		titleBuffer.append(COMRANK_DEFAULT + "\t");
		titleBuffer.append("时间（毫秒值）" + "\t");
		titleBuffer.append("明显错误信息" + "\t");
		titleBuffer.append("与电子版对比错误信息" + "\r\n");
		String titleStr = titleBuffer.toString();
		ArrayList<String> titleList = new ArrayList<>();
		titleList.add("序号");
		titleList.add("文件名");
		titleList.add(TITLE_DEFAULT);
		titleList.add(SCHOOL_DEFAULT);
		titleList.add(CLASSSTR_DEFAULT);
		titleList.add(SIGNDATE_DAFAULT);
		titleList.add(DEPARTMENTSIGNDATE_DEFAULT);
		titleList.add(SCHOOLSIGNDATE_DEFAULT);
		titleList.add(NAME_DEFAULT);
		titleList.add(RIGHTID_DEFAULT);
		titleList.add(DEPARTMENT_DEFAULT);
		titleList.add(MAJOR_DEFAULT);
		titleList.add(STUDENTID_DEFAULT);
		titleList.add(SEX_DEFAULT);
		titleList.add(NATION_DEFAULT);
		titleList.add(COMEDATE_DEFAULT);
		titleList.add(POORLEVEL_DEFAULT);
		titleList.add(STURANK_DEFAULT);
		titleList.add(TOTALNUM_DEFAULT);
		titleList.add(COMRANK_DEFAULT);
		titleList.add("时间（毫秒值）");
		titleList.add("明显错误信息");
		titleList.add("与电子版对比错误信息");
		paperList.add(titleList);
		System.out.print(titleStr);
		try {
			fw.write(titleStr);
		} catch (IOException e) {
			logger.log(Level.WARNING, "标题不能输出到txt文件！！！" + CustomExceptionTool.getExceptionMsg(e));
		}

		for (int i = 0; i < files.length; i++) {
			File file = files[i];
			// 判断是否是图片
			if (file.getName().matches(".*\\.(?i)(jpg|jpeg|gif|bmp|png)")) {
				// 标题
				String title = TITLE_DEFAULT;
				// 学校
				String school = SCHOOL_DEFAULT;
				// 学院
				String department = DEPARTMENT_DEFAULT;
				// 专业
				String major = MAJOR_DEFAULT;
				// 姓名
				String name = NAME_DEFAULT;
				// 性别
				String sex = SEX_DEFAULT;
				// 民族
				String nation = NATION_DEFAULT;
				// 入学年月
				String comeDate = COMEDATE_DEFAULT;
				// 学号
				String studentID = STUDENTID_DEFAULT;
				// 班级
				String classStr = CLASSSTR_DEFAULT;
				// 联系电话
				String phone = PHONE_DEFAULT;
				// 身份证号
				String rightID = RIGHTID_DEFAULT;
				// 家庭经济困难学生认定等级
				String poorLevel = POORLEVEL_DEFAULT;
				// 总人数
				String totalNum = TOTALNUM_DEFAULT;
				// 学习成绩排名
				String stuRank = STURANK_DEFAULT;
				// 综合考评排名(如有):
				String comRank = COMRANK_DEFAULT;
				// 签名日期
				String signDate = SIGNDATE_DAFAULT;
				// 学院盖章日期
				String departmentSignDate = DEPARTMENTSIGNDATE_DEFAULT;
				// 学校盖章日期
				String schoolSignDate = SCHOOLSIGNDATE_DEFAULT;

				// excel表中map数据
				HashMap<String, String> map = new HashMap<>();
				if (list != null && list.size() > i) {
					map = list.get(i);
				} else {
					logger.log(Level.WARNING, "电子版excel数据不足");
				}

				String path = file.getAbsolutePath();

				String jsonOriginStr = "";
				// 存放识别得到的原始信息
				JSONArray tempList = new JSONArray();
				try {
					JSONObject json = ocrClient.sampleAccurate(client, path);
					if (json.toString()
							.equals("{\"error_msg\":\"Open api daily request limit reached\",\"error_code\":17}")) {
						logger.log(Level.SEVERE, "一天识别次数已经达到上限！！！");
						System.exit(0);
					} else if (json.toString()
							.equals("{\"error_msg\":\"IAM Certification failed\",\"error_code\":14}")) {
						logger.log(Level.SEVERE, "百度云验证失败，app_id，api_key，secret_key之一填写错误！！！");
						System.exit(0);
					}
					tempList = (JSONArray) json.get("words_result");
					jsonOriginStr = tempList.toString().substring(0, tempList.toString().length() >> 1);
				} catch (Exception e) {
					e.printStackTrace();
					logger.log(Level.WARNING, "识别该图片失败或者一天识别次数已经达到上限！！！" + CustomExceptionTool.getExceptionMsg(e));
					continue;
				}
				// 输出原始的json格式字符串
				logger.info(tempList.toString());
				// 用来核对的临时map集合 (将空值设置为默认值, 检验是否有明显错误,校对是否与电子版一致)
				HashMap<String, String> checkMap = new HashMap<>();
				// 明显的错误信息
				StringBuffer errorMsg = new StringBuffer("");
				// 与电子版不同的错误信息
				StringBuffer checkErrorMsg = new StringBuffer("");
				try {

					for (int j = 0; j < tempList.length() - 2; j++) {
						// 获取当前循环的值
						JSONObject jsonID = new JSONObject(tempList.get(j).toString());
						String idNum = jsonID.get("words").toString();
						// 获取下一个循环的值
						jsonID = new JSONObject(tempList.get(j + 1).toString());
						String wordsValue = jsonID.get("words").toString();

						// 获取下两个循环的的值
						jsonID = new JSONObject(tempList.get(j + 2).toString());
						String wordsNext2 = jsonID.get("words").toString();

						// 获取标题
						if (idNum.contains("专科生国家励志奖学金申请审批表")) {
							title = idNum;
						}
						// 获取学校
						if (idNum.equals("学校:") && SCHOOL_DEFAULT.equals(school) && wordsValue.equals("广东金融学院")) {
							school = wordsValue;
						} else if (idNum.contains("学校:") && idNum.length() >= 9 && SCHOOL_DEFAULT.equals(school)) {
							school = idNum.substring(3, 9);
						}

						// 获取学院
						if (idNum.contains("院系:")) {
							department = ocrClient.getSubStr(idNum, "系:", "专业:");
							// 这里可以获取所有学院的名称，判断该名称是否在我们学院中
							// TODO
						}
						// 获取专业
						if (idNum.contains("专业:")) {
							major = ocrClient.getSubStr(idNum, "业:", NAME_DEFAULT);
						}
						// 获取姓名
						if (NAME_DEFAULT.equals(idNum)) {
							name = wordsValue;
							if (-1 != name.indexOf(SEX_DEFAULT)) {
								name = name.substring(0, name.indexOf(SEX_DEFAULT));
							}
						} else if (idNum.contains(NAME_DEFAULT) && name.equals(NAME_DEFAULT)) {
							name = ocrClient.getSubStr(idNum, NAME_DEFAULT, SEX_DEFAULT);
						}
						// 获取性别
						if (SEX_DEFAULT.equals(idNum)) {
							sex = wordsValue;
						} else if (idNum.contains(SEX_DEFAULT) && sex.equals(SEX_DEFAULT)) {
							sex = ocrClient.getSubStr(idNum, SEX_DEFAULT, NATION_DEFAULT);
						}
						// 获取民族
						if (NATION_DEFAULT.equals(idNum)) {
							nation = wordsValue;
							// 这里应该判断56个民族(有没有族字都可以)
							// TODO
						} else if (idNum.contains(NATION_DEFAULT) && nation.equals(NATION_DEFAULT)) {
							if (idNum.indexOf(NATION_DEFAULT) + 2 == idNum.length() && !wordsValue.contains("入学")) {
								nation = wordsValue;
							} else {
								nation = ocrClient.getSubStr(idNum, NATION_DEFAULT, "入学");
								if (StringUtils.isEmpty(nation)) {
									nation = wordsValue.substring(0, wordsValue.indexOf("入学"));
								}
							}
							// 这里应该判断56个民族(有没有族字都可以)
							// TODO
						}
						// 获取入学年份
						if (COMEDATE_DEFAULT.equals(idNum)) {
							comeDate = wordsValue;

						} else if (idNum.contains(COMEDATE_DEFAULT) && comeDate.equals(COMEDATE_DEFAULT)) {
							comeDate = ocrClient.getSubStr(idNum, "年月", "基本");
						}
						// 获取学号
						if ("基本学号".equals(idNum)) {
							studentID = wordsValue;
							if (-1 != studentID.indexOf(CLASSSTR_DEFAULT)) {
								studentID = studentID.substring(0, studentID.indexOf(CLASSSTR_DEFAULT));
							}
						} else if (STUDENTID_DEFAULT.equals(studentID) && idNum.contains(STUDENTID_DEFAULT)) {
							// 获取后的学号
							studentID = ocrClient.cutByRegex(idNum, "\\D+(\\d{9})\\D*");
						}
						// 获取班级
						if (CLASSSTR_DEFAULT.equals(idNum) && wordsValue.matches("\\d{7}")) {
							classStr = wordsValue;
						} else if (idNum.contains(CLASSSTR_DEFAULT) && classStr.equals(CLASSSTR_DEFAULT)) {
							classStr = ocrClient.getSubStr(idNum, CLASSSTR_DEFAULT, "联系");
						} else if (CLASSSTR_DEFAULT.equals(idNum)) {
							classStr = ocrClient.cutByRegex(jsonOriginStr.substring(140, jsonOriginStr.length()),
									"\\D*[班别]\\D*(\\d{0,9})\\D*[联系电话]");
						}
						// 获取联系电话
						if (PHONE_DEFAULT.equals(idNum)) {
							phone = wordsValue;
						} else if (idNum.contains(PHONE_DEFAULT) && phone.equals(PHONE_DEFAULT)) {
							phone = ocrClient.getSubStr(idNum, "电话", "情况");
						}

						// 获取身份证号
						if (("情况身份证号".equals(idNum) || RIGHTID_DEFAULT.equals(idNum))) {
							rightID = wordsValue;
						} else if (RIGHTID_DEFAULT.equals(rightID) && idNum.contains(RIGHTID_DEFAULT)) {
							// 获取身份证号
							rightID = ocrClient.cutByRegex(idNum, ".*[身份证号].*(\\d{17}[\\d|x|X]{1}).*");
						}

						// 获取家庭经济困难学生认定等级
						if (idNum.equals("口一般困难口比较困难特殊困难") || idNum.equals("口一般困难口比较闲难特殊困难")
								|| idNum.equals("家庭经济困难学生认定等级口一般困难口比较困难特殊困难")) {
							poorLevel = "特殊困难";
						} else if (idNum.equals("口一般困难比较困难口特殊困难") || idNum.equals("家庭经济困难学生认定等级口一般困难比较困难口特殊困难")) {
							poorLevel = "比较困难";
						} else if (idNum.equals("一般困难口比较困难口特殊困难") || idNum.equals("家庭经济困难学生认定等级一般困难口比较困难口特殊困难")) {
							poorLevel = "一般困难";
						} else if (idNum.equals("口一般困难口比较困难口特殊困难") || idNum.equals("一般困难比较困难口特殊困难")
								|| idNum.equals("口一般困难比较困难特殊困难") || idNum.equals("一般困难比较困难特殊困难")
								|| idNum.equals("一般困难口比较困难特殊困难") || idNum.equals("家庭经济困难学生认定等级口一般困难口比较困难口特殊困难")
								|| idNum.equals("家庭经济困难学生认定等级一般困难比较困难口特殊困难")
								|| idNum.equals("家庭经济困难学生认定等级口一般困难比较困难特殊困难") || idNum.equals("家庭经济困难学生认定等级一般困难比较困难特殊困难")
								|| idNum.equals("家庭经济困难学生认定等级一般困难口比较困难特殊困难")) {
							errorMsg.append("家庭经济困难学生认定等级可能勾选错误；");
						}

						// 获取总人数
						if ((idNum.equals("总人数:") || "学习总人数".equals(idNum)) && totalNum.equals(TOTALNUM_DEFAULT)) {
							totalNum = wordsValue;
							if (wordsValue.matches("\\d+\\D+")) {
								totalNum = ocrClient.cutByRegex(wordsValue, "(\\d+)\\D+");
							}
						} else if (idNum.contains(TOTALNUM_DEFAULT) && totalNum.equals(TOTALNUM_DEFAULT)) {
							// 数字匹配
							Matcher matcher = Pattern.compile("\\d+").matcher(idNum);
							while (matcher.find()) {
								totalNum = matcher.group();
							}
						}

						// 获取学习成绩排名
						if (idNum.contains("习成绩排名") && STURANK_DEFAULT.equals(stuRank)) {
							// 如果字符太长，会查找不到
							stuRank = ocrClient.cutByRegex(jsonOriginStr.substring(280, jsonOriginStr.length()),
									".*[习成绩排名]\\D*(\\d{0,3})\\D*[合考评].*");
							if (StringUtils.isEmpty(stuRank)) {
								stuRank = ocrClient.cutByRegex(jsonOriginStr.substring(280, jsonOriginStr.length()),
										".*[习成绩排名]\\D*(\\d{0,3})\\D*[修课].*");
							}
						}

						// 获取综合测评排名
						if ((idNum.equals(";综合考评排名(如有):") || idNum.equals("综合考评排名(如有):")
								|| idNum.equals(":综合考评排名(如有):")) && wordsValue.matches("\\d+")
								&& COMRANK_DEFAULT.equals(comRank)) {
							comRank = wordsValue;
						} else if (COMRANK_DEFAULT.equals(comRank)) {
							comRank = ocrClient.cutByRegex(jsonOriginStr.substring(160, jsonOriginStr.length()),
									".*[考评排名]\\D*(\\d{0,3})\\D*[修课]{1}.*");
						}

						// 获取个人签名日期
						if (SIGNDATE_DAFAULT.equals(signDate)) {
							signDate = ocrClient.cutByRegex(
									tempList.toString().substring(tempList.toString().length() >> 1,
											tempList.toString().length()),
									".*[请人签名].*(\\d{4}[年]\\d{0,2}[月]\\d{0,2}[日]).*[推荐].*");
						}
						// 获取学院签名日期
						if (DEPARTMENTSIGNDATE_DEFAULT.equals(departmentSignDate)) {
							departmentSignDate = ocrClient.cutByRegex(
									tempList.toString().substring(tempList.toString().length() >> 1,
											tempList.toString().length()),
									".*[意推荐].*(\\d{4}[年]\\d{0,2}[月]\\d{0,2}[日]).*[评审].*");
						}
						// 获取学校签名日期
						if (SCHOOLSIGNDATE_DEFAULT.equals(schoolSignDate)) {
							schoolSignDate = ocrClient.cutByRegex(
									tempList.toString().substring(tempList.toString().length() >> 1,
											tempList.toString().length()),
									".*[经评审].*(\\d{4}[年]\\d{0,2}[月]\\d{0,2}[日]).*[东省].*");
						}

						// 检查是否含有书信体
						if (idNum.contains("尊敬的")) {
							errorMsg.append("可能申请理由含有书信体；");
						}
					}

					checkMap.put(TITLE_DEFAULT, title);
					checkMap.put(SCHOOL_DEFAULT, school);
					checkMap.put(DEPARTMENT_DEFAULT, department);
					checkMap.put(MAJOR_DEFAULT, major);
					checkMap.put(NAME_DEFAULT, name);
					checkMap.put(SEX_DEFAULT, sex);
					checkMap.put(NATION_DEFAULT, nation);
					checkMap.put(COMEDATE_DEFAULT, comeDate);
					checkMap.put(STUDENTID_DEFAULT, studentID);
					checkMap.put(CLASSSTR_DEFAULT, classStr);
					checkMap.put(PHONE_DEFAULT, phone);
					checkMap.put(RIGHTID_DEFAULT, rightID);
					checkMap.put(POORLEVEL_DEFAULT, poorLevel);
					checkMap.put(TOTALNUM_DEFAULT, totalNum);
					checkMap.put(STURANK_DEFAULT, stuRank);
					checkMap.put(COMRANK_DEFAULT, comRank);
					checkMap.put(SIGNDATE_DAFAULT, signDate);
					checkMap.put(DEPARTMENTSIGNDATE_DEFAULT, departmentSignDate);
					checkMap.put(SCHOOLSIGNDATE_DEFAULT, schoolSignDate);

					// 将空值设置为默认值
					ocrClient.setNullToDefault(checkMap);

					// 检验是否有明显错误
					ocrClient.existObviousErr(checkMap, errorMsg);

					// 校对是否与电子版一致
					ocrClient.equalsWithElec(checkMap, map, checkErrorMsg);

				} catch (Exception e) {
					logger.log(Level.WARNING,
							"获取识别内容出现错误！" + CustomExceptionTool.getExceptionMsg(e) + "\n" + "识别得到的原始值：" + tempList);
				}

				StringBuffer buffer = new StringBuffer("");
				num = num + 1;
				buffer.append(num + "\t" + file.getName() + "\t");
				buffer.append(checkMap.get(TITLE_DEFAULT) + "\t");
				buffer.append(checkMap.get(SCHOOL_DEFAULT) + "\t");
				buffer.append(checkMap.get(CLASSSTR_DEFAULT) + "\t");
				buffer.append(checkMap.get(SIGNDATE_DAFAULT) + "\t");
				buffer.append(checkMap.get(DEPARTMENTSIGNDATE_DEFAULT) + "\t");
				buffer.append(checkMap.get(SCHOOLSIGNDATE_DEFAULT) + "\t");
				buffer.append(checkMap.get(NAME_DEFAULT) + "\t");
				buffer.append(checkMap.get(RIGHTID_DEFAULT) + "\t");
				buffer.append(checkMap.get(DEPARTMENT_DEFAULT) + "\t");
				buffer.append(checkMap.get(MAJOR_DEFAULT) + "\t");
				buffer.append(checkMap.get(STUDENTID_DEFAULT) + "\t");
				buffer.append(checkMap.get(SEX_DEFAULT) + "\t");
				buffer.append(checkMap.get(NATION_DEFAULT) + "\t");
				buffer.append(checkMap.get(COMEDATE_DEFAULT) + "\t");
				buffer.append(checkMap.get(POORLEVEL_DEFAULT) + "\t");
				buffer.append(checkMap.get(STURANK_DEFAULT) + "\t");
				buffer.append(checkMap.get(TOTALNUM_DEFAULT) + "\t");
				buffer.append(checkMap.get(COMRANK_DEFAULT) + "\t");
				buffer.append(System.currentTimeMillis() + "\t");
				buffer.append(errorMsg + "\t");
				buffer.append(checkErrorMsg + "\r\n");
				String studentInfoStr = buffer.toString();
				ArrayList<String> rowList = new ArrayList<>();
				rowList.add(String.valueOf(num));
				rowList.add(file.getName());
				rowList.add(checkMap.get(TITLE_DEFAULT));
				rowList.add(checkMap.get(SCHOOL_DEFAULT));
				rowList.add(checkMap.get(CLASSSTR_DEFAULT));
				rowList.add(checkMap.get(SIGNDATE_DAFAULT));
				rowList.add(checkMap.get(DEPARTMENTSIGNDATE_DEFAULT));
				rowList.add(checkMap.get(SCHOOLSIGNDATE_DEFAULT));
				rowList.add(checkMap.get(NAME_DEFAULT));
				rowList.add(checkMap.get(RIGHTID_DEFAULT));
				rowList.add(checkMap.get(DEPARTMENT_DEFAULT));
				rowList.add(checkMap.get(MAJOR_DEFAULT));
				rowList.add(checkMap.get(STUDENTID_DEFAULT));
				rowList.add(checkMap.get(SEX_DEFAULT));
				rowList.add(checkMap.get(NATION_DEFAULT));
				rowList.add(checkMap.get(COMEDATE_DEFAULT));
				rowList.add(checkMap.get(POORLEVEL_DEFAULT));
				rowList.add(checkMap.get(STURANK_DEFAULT));
				rowList.add(checkMap.get(TOTALNUM_DEFAULT));
				rowList.add(checkMap.get(COMRANK_DEFAULT));
				rowList.add(String.valueOf(System.currentTimeMillis()));
				rowList.add(errorMsg.toString());
				rowList.add(checkErrorMsg.toString());
				paperList.add(rowList);

				System.out.print(studentInfoStr);
				try {
					fw.write(studentInfoStr);
					fw.flush();
				} catch (IOException e) {
					logger.log(Level.WARNING, "输出到txt文件中错误！！" + CustomExceptionTool.getExceptionMsg(e));
				}
				if (!org.apache.commons.lang.StringUtils.isEmpty(errorMsg.toString())) {
					errorMap.put(String.valueOf(num), errorMsg.toString());
				}
				if (!org.apache.commons.lang.StringUtils.isEmpty(checkErrorMsg.toString())) {
					checkErrorMap.put(String.valueOf(num), checkErrorMsg.toString());
				}
			} else {
				logger.log(Level.SEVERE, "要识别文件错误，只能识别图片！！！");
			}

		}

		try {
			// 输出到excel表中
			option.writeExcel(paperList, targetExcelPath);
			// 输出到txt文件中
			Set<String> set = errorMap.keySet();
			for (String key : set) {
				String errMsg = key + "\t" + "明显的错误信息：" + errorMap.get(key) + "\r\n";
				System.out.print(errMsg);
				fw.write(errMsg);
				fw.flush();
			}
			// 输出核对错误信息
			Set<String> checkSet = checkErrorMap.keySet();
			for (String key : checkSet) {
				String errMsg = key + "\t" + "与电子表核对的错误信息：" + checkErrorMap.get(key) + "\r\n";
				System.out.print(errMsg);
				fw.write(errMsg);
				fw.flush();
			}
			fw.close();
			System.out.println("已经全部输出到" + targetExcelPath + "文件中了！");
		} catch (Exception e) {
			e.printStackTrace();
			logger.log(Level.WARNING, "输出识别内容信息的时候发生错误！！" + CustomExceptionTool.getExceptionMsg(e));
		}

	}

	/**
	 * 通用文字识别 调用例子
	 *
	 * @param client
	 */
	public JSONObject sample(AipOcr client, String path) {
		// 传入可选参数调用接口
		HashMap<String, String> options = new HashMap<>();
		options.put("detect_direction", "false");
		options.put("probability", "false");

		// 参数为本地图片路径
		String image = path;
		JSONObject res = client.basicGeneral(image, options);
		// System.out.println(res.toString(2));
		return res;
	}

	/**
	 * 通用文字识别(高精度版) 调用例子
	 *
	 * @param client
	 */
	public JSONObject sampleAccurate(AipOcr client, String path) {
		// 传入可选参数调用接口
		HashMap<String, String> options = new HashMap<>();
		options.put("detect_direction", "false");
		options.put("probability", "false");

		// 参数为本地图片路径
		String image = path;
		JSONObject res = client.basicAccurateGeneral(image, options);
		// System.out.println(res.toString(2));
		return res;
	}

	/**
	 * 表格文字识别方法
	 *
	 * @param client
	 * @param path
	 */
	public void tableRecognition(AipOcr client, String path) {
		// 异步接口

		byte[] file = image2byte(path);
		// 使用封装的同步轮询接口
		JSONObject jsonres = client.tableRecognizeToJson(file, 20000);
		System.out.println(jsonres.toString(2));
	}

	/**
	 * 将图片转换为byte数组 
	 *
	 * @param path
	 * @return
	 */
	private byte[] image2byte(String path) {
		byte[] data = null;
		FileImageInputStream input = null;
		try {
			input = new FileImageInputStream(new File(path));
			ByteArrayOutputStream output = new ByteArrayOutputStream();
			byte[] buf = new byte[1024];
			int numBytesRead = 0;
			while ((numBytesRead = input.read(buf)) != -1) {
				output.write(buf, 0, numBytesRead);
			}
			data = output.toByteArray();
			output.close();
			input.close();
		} catch (Exception e) {
			e.printStackTrace();
			logger.log(Level.SEVERE, "将图片转换为byte数组错误！！" + CustomExceptionTool.getExceptionMsg(e));
		}
		return data;
	}

	/**
	 * 根据表达式，从某字符串中提取自己想要的子字符串
	 *
	 * @param str
	 * @param regex
	 * @return
	 */
	protected String cutByRegex(String str, String regex) {

		String reg = regex;
		String s = str;
		Pattern p2 = Pattern.compile(reg);
		Matcher m2 = p2.matcher(s);
		if (m2.find()) {
			String subStr = m2.group(1);
			return subStr;// 组提取字符串
		}
		return null;

	}

	/**
	 * 根据某首字符和尾字符，截取字符串
	 *
	 * @param str
	 *            要截取的字符串（只能是两位）
	 * @param headIndexStr
	 *            截取后的首字符（只能是两位）
	 * @param lastIndexStr
	 *            尾字符
	 * @return 如果首字母不存在，返回空
	 */
	protected String getSubStr(String str, String headIndexStr, String lastIndexStr) {
		int index = str.indexOf(headIndexStr);
		if (index == -1) {
			return null;
		}
		int lastIndex = str.indexOf(lastIndexStr);
		if (lastIndex == -1) {
			lastIndex = str.length();
		}
		return str.substring(index + 2, lastIndex);
	}

	/**
	 * 将空的值设置为原来的默认值
	 *
	 * @param map
	 *            参数
	 */
	private void setNullToDefault(Map<String, String> map) {
		// 检查是否为空，如果是空，将其设置为默认值
		if (StringUtils.isEmpty(map.get(TITLE_DEFAULT))) {
			map.put(TITLE_DEFAULT, TITLE_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(SCHOOL_DEFAULT))) {
			map.put(SCHOOL_DEFAULT, SCHOOL_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(DEPARTMENT_DEFAULT))) {
			map.put(DEPARTMENT_DEFAULT, DEPARTMENT_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(MAJOR_DEFAULT))) {
			map.put(MAJOR_DEFAULT, MAJOR_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(NAME_DEFAULT))) {
			map.put(NAME_DEFAULT, NAME_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(SEX_DEFAULT))) {
			map.put(SEX_DEFAULT, SEX_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(NATION_DEFAULT))) {
			map.put(NATION_DEFAULT, NATION_DEFAULT);
		}

		if (StringUtils.isEmpty(map.get(COMEDATE_DEFAULT))) {
			map.put(COMEDATE_DEFAULT, COMEDATE_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(STUDENTID_DEFAULT))) {
			map.put(STUDENTID_DEFAULT, STUDENTID_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(CLASSSTR_DEFAULT))) {
			map.put(CLASSSTR_DEFAULT, CLASSSTR_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(PHONE_DEFAULT))) {
			map.put(PHONE_DEFAULT, PHONE_DEFAULT);
		}

		if (StringUtils.isEmpty(map.get(RIGHTID_DEFAULT))) {
			map.put(RIGHTID_DEFAULT, RIGHTID_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(POORLEVEL_DEFAULT))) {
			map.put(POORLEVEL_DEFAULT, POORLEVEL_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(TOTALNUM_DEFAULT))) {
			map.put(TOTALNUM_DEFAULT, TOTALNUM_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(STURANK_DEFAULT))) {
			map.put(STURANK_DEFAULT, STURANK_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(COMRANK_DEFAULT))) {
			map.put(COMRANK_DEFAULT, COMRANK_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(SIGNDATE_DAFAULT))) {
			map.put(SIGNDATE_DAFAULT, SIGNDATE_DAFAULT);
		}
		if (StringUtils.isEmpty(map.get(DEPARTMENTSIGNDATE_DEFAULT))) {
			map.put(DEPARTMENTSIGNDATE_DEFAULT, DEPARTMENTSIGNDATE_DEFAULT);
		}
		if (StringUtils.isEmpty(map.get(SCHOOLSIGNDATE_DEFAULT))) {
			map.put(SCHOOLSIGNDATE_DEFAULT, SCHOOLSIGNDATE_DEFAULT);
		}

	}

	/**
	 * 检查是否存在明显错误
	 *
	 * @param map
	 * @param errorMsg
	 * @return
	 */
	private StringBuffer existObviousErr(Map<String, String> map, StringBuffer errorMsg) throws Exception {
		// 检查标题是否错误
		if (!map.get(TITLE_DEFAULT).equals("(2017-2018学年)本专科生国家励志奖学金申请审批表")) {
			errorMsg.append("标题错误；");
		}
		if (!map.get(SCHOOL_DEFAULT).equals("广东金融学院")) {
			errorMsg.append("学校写错；");
		}
		if ((!"女".equals(map.get(SEX_DEFAULT))) && (!"男".equals(map.get(SEX_DEFAULT)))) {
			errorMsg.append("性别填写错误；");
		}
		try {
			if (!map.get(COMEDATE_DEFAULT).substring(0, 4).matches("\\d{4}")) {
				errorMsg.append("入学年月错误；");
			} else {
				int year = Integer.parseInt(map.get(COMEDATE_DEFAULT).substring(0, 4));
				int nowYear = Calendar.getInstance().get(Calendar.YEAR);
				if ((year > nowYear - 1 || year < nowYear - 3)
						|| (!"9".equals(map.get(COMEDATE_DEFAULT).substring(5, 6)))) {
					errorMsg.append("入学年月错误；");
				}
			}
		} catch (Exception e) {
			logger.log(Level.WARNING, "转换入学年月错误！！" + CustomExceptionTool.getExceptionMsg(e));
			errorMsg.append("入学年月错误；");
		}
		if (!map.get(STUDENTID_DEFAULT).matches("[1]+\\d{8}")) {
			errorMsg.append("学号错误；");
		}
		if (!map.get(CLASSSTR_DEFAULT).matches("\\d{7}") || (map.get(STUDENTID_DEFAULT).matches("[1]+\\d{8}")
				&& !map.get(CLASSSTR_DEFAULT).equals(map.get(STUDENTID_DEFAULT).substring(0, 7)))) {
			errorMsg.append("班级错误；");
		}
		if (!map.get(PHONE_DEFAULT).matches("\\d{11}")) {
			errorMsg.append("联系电话错误");
		}
		if (!map.get(RIGHTID_DEFAULT).matches("\\d{17}([0-9]|x|X)")) {
			errorMsg.append("身份证号错误；");
		}
		if (!map.get(TOTALNUM_DEFAULT).matches("\\d+") || Integer.parseInt(map.get(TOTALNUM_DEFAULT)) > 100) {
			errorMsg.append("学习总人数可能错误；");
		}
		if ((!map.get(STURANK_DEFAULT).matches("\\d+")) || (Integer.parseInt(map.get(STURANK_DEFAULT)) >= 30)) {
			errorMsg.append("学习成绩可能错误；");
		}
		if ((!map.get(COMRANK_DEFAULT).matches("\\d+")) || (Integer.parseInt(map.get(COMRANK_DEFAULT)) >= 30)) {
			errorMsg.append("综合考评排名可能错误；");
		}
		if ((!map.get(SIGNDATE_DAFAULT).matches("\\d{4}[年]\\d{1,2}[月]\\d{1,2}[日]"))) {
			errorMsg.append("个人签名日期可能错误；");
		}
		if ((!map.get(DEPARTMENTSIGNDATE_DEFAULT).matches("\\d{4}[年]\\d{1,2}[月]\\d{1,2}[日]"))) {
			errorMsg.append("学院盖章日期可能错误；");
		}
		if ((!map.get(SCHOOLSIGNDATE_DEFAULT).matches("\\d{4}[年]\\d{1,2}[月]\\d{1,2}[日]"))) {
			errorMsg.append("学校盖章日期可能错误；");
		}
		return errorMsg;
	}

	/**
	 * 检查纸质版和电子版的是否一样
	 *
	 * @param paperMap
	 *            纸质版信息
	 * @param elecMap
	 *            电子版信息
	 * @param checkErrorMsg
	 *            核对电子版的错误信息
	 * @return 一样返回ture，不一样返回false
	 */
	private Boolean equalsWithElec(Map<String, String> paperMap, Map<String, String> elecMap,
			StringBuffer checkErrorMsg) {
		if (elecMap.containsKey(TITLE_DEFAULT) && (!paperMap.get(TITLE_DEFAULT).equals(elecMap.get(TITLE_DEFAULT)))) {
			checkErrorMsg.append(TITLE_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(TITLE_DEFAULT) + "，纸质版为"
					+ paperMap.get(TITLE_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(SCHOOL_DEFAULT)
				&& (!paperMap.get(SCHOOL_DEFAULT).equals(elecMap.get(SCHOOL_DEFAULT)))) {
			checkErrorMsg.append(SCHOOL_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(SCHOOL_DEFAULT) + "，纸质版为"
					+ paperMap.get(SCHOOL_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(DEPARTMENT_DEFAULT)
				&& (!paperMap.get(DEPARTMENT_DEFAULT).equals(elecMap.get(DEPARTMENT_DEFAULT)))) {
			checkErrorMsg.append(DEPARTMENT_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(DEPARTMENT_DEFAULT) + "，纸质版为"
					+ paperMap.get(DEPARTMENT_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(MAJOR_DEFAULT) && (!paperMap.get(MAJOR_DEFAULT).equals(elecMap.get(MAJOR_DEFAULT)))) {
			checkErrorMsg.append(MAJOR_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(MAJOR_DEFAULT) + "，纸质版为"
					+ paperMap.get(MAJOR_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(NAME_DEFAULT) && (!paperMap.get(NAME_DEFAULT).equals(elecMap.get(NAME_DEFAULT)))) {
			checkErrorMsg.append(NAME_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(NAME_DEFAULT) + "，纸质版为"
					+ paperMap.get(NAME_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(SEX_DEFAULT) && (!paperMap.get(SEX_DEFAULT).equals(elecMap.get(SEX_DEFAULT)))) {
			checkErrorMsg.append(SEX_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(SEX_DEFAULT) + "，纸质版为"
					+ paperMap.get(SEX_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(NATION_DEFAULT)
				&& (!paperMap.get(NATION_DEFAULT).equals(elecMap.get(NATION_DEFAULT)))) {
			checkErrorMsg.append(NATION_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(NATION_DEFAULT) + "，纸质版为"
					+ paperMap.get(NATION_DEFAULT) + "）");
		}
		if (elecMap.containsKey(COMEDATE_DEFAULT)
				&& (!elecMap.get(COMEDATE_DEFAULT).equals(paperMap.get(COMEDATE_DEFAULT)))) {
			checkErrorMsg.append(COMEDATE_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(COMEDATE_DEFAULT) + "，纸质版为"
					+ paperMap.get(COMEDATE_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(STUDENTID_DEFAULT)
				&& (!elecMap.get(STUDENTID_DEFAULT).equals(paperMap.get(STUDENTID_DEFAULT)))) {
			checkErrorMsg.append(STUDENTID_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(STUDENTID_DEFAULT) + "，纸质版为"
					+ paperMap.get(STUDENTID_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(CLASSSTR_DEFAULT)
				&& (!elecMap.get(CLASSSTR_DEFAULT).equals(paperMap.get(CLASSSTR_DEFAULT)))) {
			checkErrorMsg.append(CLASSSTR_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(CLASSSTR_DEFAULT) + "，纸质版为"
					+ paperMap.get(CLASSSTR_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(PHONE_DEFAULT) && (!paperMap.get(PHONE_DEFAULT).equals(elecMap.get(PHONE_DEFAULT)))) {
			checkErrorMsg.append(PHONE_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(PHONE_DEFAULT) + "，纸质版为"
					+ paperMap.get(PHONE_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(RIGHTID_DEFAULT)
				&& (!paperMap.get(RIGHTID_DEFAULT).equalsIgnoreCase(elecMap.get(RIGHTID_DEFAULT)))) {
			checkErrorMsg.append(RIGHTID_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(RIGHTID_DEFAULT) + "，纸质版为"
					+ paperMap.get(RIGHTID_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(POORLEVEL_DEFAULT)
				&& (!paperMap.get(POORLEVEL_DEFAULT).equals(elecMap.get(POORLEVEL_DEFAULT)))) {
			checkErrorMsg.append(POORLEVEL_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(POORLEVEL_DEFAULT) + "，纸质版为"
					+ paperMap.get(POORLEVEL_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(TOTALNUM_DEFAULT)
				&& (!paperMap.get(TOTALNUM_DEFAULT).equals(elecMap.get(TOTALNUM_DEFAULT)))) {
			checkErrorMsg.append(TOTALNUM_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(TOTALNUM_DEFAULT) + "，纸质版为"
					+ paperMap.get(TOTALNUM_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(STURANK_DEFAULT)
				&& (!paperMap.get(STURANK_DEFAULT).equals(elecMap.get(STURANK_DEFAULT)))) {
			checkErrorMsg.append(STURANK_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(STURANK_DEFAULT) + "，纸质版为"
					+ paperMap.get(STURANK_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(COMRANK_DEFAULT)
				&& (!paperMap.get(COMRANK_DEFAULT).equals(elecMap.get(COMRANK_DEFAULT)))) {
			checkErrorMsg.append(COMRANK_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(COMRANK_DEFAULT) + "，纸质版为"
					+ paperMap.get(COMRANK_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(SIGNDATE_DAFAULT)
				&& (!paperMap.get(SIGNDATE_DAFAULT).equals(elecMap.get(SIGNDATE_DAFAULT)))) {
			checkErrorMsg.append(SIGNDATE_DAFAULT + "与电子版不一致（电子版为" + elecMap.get(SIGNDATE_DAFAULT) + "，纸质版为"
					+ paperMap.get(SIGNDATE_DAFAULT) + "）；");
		}
		if (elecMap.containsKey(DEPARTMENTSIGNDATE_DEFAULT)
				&& (!paperMap.get(DEPARTMENTSIGNDATE_DEFAULT).equals(elecMap.get(DEPARTMENTSIGNDATE_DEFAULT)))) {
			checkErrorMsg.append(DEPARTMENTSIGNDATE_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(DEPARTMENTSIGNDATE_DEFAULT)
					+ "，纸质版为" + paperMap.get(DEPARTMENTSIGNDATE_DEFAULT) + "）；");
		}
		if (elecMap.containsKey(SCHOOLSIGNDATE_DEFAULT)
				&& (!paperMap.get(SCHOOLSIGNDATE_DEFAULT).equals(elecMap.get(SCHOOLSIGNDATE_DEFAULT)))) {
			checkErrorMsg.append(SCHOOLSIGNDATE_DEFAULT + "与电子版不一致（电子版为" + elecMap.get(SCHOOLSIGNDATE_DEFAULT) + "，纸质版为"
					+ paperMap.get(SCHOOLSIGNDATE_DEFAULT) + "）；");
		}

		if (StringUtils.isEmpty(checkErrorMsg.toString())) {
			return true;
		}
		return false;
	}

}