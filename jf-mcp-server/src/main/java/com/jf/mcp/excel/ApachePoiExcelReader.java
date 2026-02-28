package com.jf.mcp.excel;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

@Component
public class ApachePoiExcelReader
{
	private final ObjectMapper objectMapper = new ObjectMapper();

	/**
	 * 读取指定工作表的所有数据，返回结构化 JSON。
	 * JSON 格式：[{"column1": value1, "column2": value2}, ...]
	 * @param filePath Excel 文件路径
	 * @param sheetName 工作表名称
	 * @param headerRow 表头行号（0-based，默认为 0）
	 * @return JSON 字符串
	 */
	public String readSheetAsJson(String filePath, String sheetName, int headerRow) throws IOException
	{
		try (FileInputStream fis = new FileInputStream(filePath);
			 Workbook workbook = new XSSFWorkbook(fis))
		{  // 支持 .xlsx；如果需 .xls，用 HSSFWorkbook
			Sheet sheet = workbook.getSheet(sheetName);
			if (sheet == null)
			{
				throw new IllegalArgumentException("Sheet '" + sheetName + "' not found");
			}

			ArrayNode jsonArray = objectMapper.createArrayNode();
			Row header = sheet.getRow(headerRow);
			if (header == null)
			{
				return jsonArray.toString();  // 空表
			}

			// 获取表头
			List<String> headers = new ArrayList<>();
			for (Cell cell : header)
			{
				headers.add(getCellValueAsString(cell));
			}

			// 读取数据行
			for (int i = headerRow + 1; i <= sheet.getLastRowNum(); i++)
			{
				Row row = sheet.getRow(i);
				if (row == null) {continue;}

				ObjectNode jsonObject = objectMapper.createObjectNode();
				for (int j = 0; j < headers.size(); j++)
				{
					Cell cell = row.getCell(j);
					jsonObject.put(headers.get(j), getCellValueAsString(cell));
				}
				jsonArray.add(jsonObject);
			}

			return jsonArray.toString();
		}
	}

	/**
	 * 精确统计 Excel 中某个关键词/数值出现的次数（支持正则）。
	 * 遍历整个 sheet 的所有单元格。
	 * @param filePath Excel 文件路径
	 * @param sheetName 工作表名称
	 * @param keyword 关键词
	 * @param useRegex 是否使用正则匹配
	 * @return 出现次数
	 */
	public int countOccurrences(String filePath, String sheetName, String keyword, boolean useRegex) throws IOException
	{
		try (FileInputStream fis = new FileInputStream(filePath);
			 Workbook workbook = new XSSFWorkbook(fis))
		{
			Sheet sheet = workbook.getSheet(sheetName);
			if (sheet == null)
			{
				throw new IllegalArgumentException("Sheet '" + sheetName + "' not found");
			}

			int count = 0;
			Pattern pattern = useRegex ? Pattern.compile(keyword) : null;

			for (Row row : sheet)
			{
				for (Cell cell : row)
				{
					String cellValue = getCellValueAsString(cell);
					if (useRegex && pattern != null)
					{
						Matcher matcher = pattern.matcher(cellValue);
						while (matcher.find())
						{
							count++;
						}
					}
					else if (cellValue.contains(keyword))
					{
						count++;
					}
				}
			}
			return count;
		}
	}

	/**
	 * 按条件过滤 Excel 行，返回匹配数据。
	 * 匹配规则：指定列的值等于 value（忽略大小写）。
	 * 返回 List<Map<String, Object>>，每行一个 Map。
	 * @param filePath Excel 文件路径
	 * @param sheetName 工作表名称
	 * @param column 列名（假设第 0 行为表头）
	 * @param value 匹配值
	 * @return 过滤后的行数据
	 */
	public List<Map<String, Object>> filterRows(String filePath, String sheetName, String column,
			String value) throws IOException
	{
		try (FileInputStream fis = new FileInputStream(filePath);
			 Workbook workbook = new XSSFWorkbook(fis))
		{
			Sheet sheet = workbook.getSheet(sheetName);
			if (sheet == null)
			{
				throw new IllegalArgumentException("Sheet '" + sheetName + "' not found");
			}

			List<Map<String, Object>> filteredRows = new ArrayList<>();

			// 获取表头和列索引
			Row header = sheet.getRow(0);
			int columnIndex = -1;
			Map<Integer, String> headerMap = new HashMap<>();
			for (Cell cell : header)
			{
				String headerValue = getCellValueAsString(cell);
				headerMap.put(cell.getColumnIndex(), headerValue);
				if (headerValue.equalsIgnoreCase(column))
				{
					columnIndex = cell.getColumnIndex();
				}
			}
			if (columnIndex == -1)
			{
				throw new IllegalArgumentException("Column '" + column + "' not found");
			}

			// 过滤数据行
			for (int i = 1; i <= sheet.getLastRowNum(); i++)
			{
				Row row = sheet.getRow(i);
				if (row == null) {continue;}

				Cell cell = row.getCell(columnIndex);
				if (getCellValueAsString(cell).equalsIgnoreCase(value))
				{
					Map<String, Object> rowMap = new LinkedHashMap<>();
					for (Cell dataCell : row)
					{
						String key = headerMap.get(dataCell.getColumnIndex());
						if (key != null)
						{
							rowMap.put(key, getCellValueAsObject(dataCell));
						}
					}
					filteredRows.add(rowMap);
				}
			}
			return filteredRows;
		}
	}

	/**
	 * 读取 Excel 元数据：工作表列表、每个 sheet 的列名、行数等。
	 * 返回 JSON 字符串：{"sheets": [{"name": "Sheet1", "rows": 10, "columns": ["Col1", "Col2"]}, ...]}
	 * @param filePath Excel 文件路径
	 * @return JSON 字符串
	 */
	public String getExcelMetadata(String filePath) throws IOException
	{
		try (FileInputStream fis = new FileInputStream(filePath);
			 Workbook workbook = new XSSFWorkbook(fis))
		{
			ObjectNode jsonRoot = objectMapper.createObjectNode();
			ArrayNode sheetsArray = objectMapper.createArrayNode();

			for (int i = 0; i < workbook.getNumberOfSheets(); i++)
			{
				Sheet sheet = workbook.getSheetAt(i);
				ObjectNode sheetNode = objectMapper.createObjectNode();
				sheetNode.put("name", sheet.getSheetName());
				sheetNode.put("rows", sheet.getLastRowNum() + 1);  // 包括表头

				// 获取列名（假设第 0 行为表头）
				ArrayNode columnsArray = objectMapper.createArrayNode();
				Row header = sheet.getRow(0);
				if (header != null)
				{
					for (Cell cell : header)
					{
						columnsArray.add(getCellValueAsString(cell));
					}
				}
				sheetNode.set("columns", columnsArray);

				sheetsArray.add(sheetNode);
			}

			jsonRoot.set("sheets", sheetsArray);
			return jsonRoot.toString();
		}
	}

	/**
	 * 读取 Excel 指定工作表中指定列的所有数据（从数据行开始，不包括表头）。
	 * 假设表头在第 0 行。
	 * @param filePath Excel 文件路径
	 * @param sheetName 工作表名称（如果 null 或空，默认第一个工作表）
	 * @param columnName 列名（字符串，必填，忽略大小写匹配）
	 * @return 该列的所有数据，作为 List<Object>（支持字符串、数字、布尔等类型）
	 */
	public List<Object> readColumnData(String filePath, String sheetName, String columnName) throws IOException
	{
		try (FileInputStream fis = new FileInputStream(filePath);
			 Workbook workbook = new XSSFWorkbook(fis))
		{
			Sheet sheet;
			if (sheetName == null || sheetName.isEmpty())
			{
				sheet = workbook.getSheetAt(0);  // 默认第一个 Sheet
			}
			else
			{
				sheet = workbook.getSheet(sheetName);
				if (sheet == null)
				{
					throw new IllegalArgumentException("Sheet '" + sheetName + "' not found");
				}
			}

			List<Object> columnData = new ArrayList<>();

			// 获取表头行（假设第 0 行）
			Row header = sheet.getRow(0);
			if (header == null)
			{
				return columnData;  // 空表，返回空列表
			}

			// 找到列索引（忽略大小写）
			int columnIndex = -1;
			for (Cell cell : header)
			{
				if (getCellValueAsString(cell).equalsIgnoreCase(columnName))
				{
					columnIndex = cell.getColumnIndex();
					break;
				}
			}
			if (columnIndex == -1)
			{
				throw new IllegalArgumentException("Column '" + columnName + "' not found in header");
			}

			// 读取数据行（从第 1 行开始）
			for (int i = 1; i <= sheet.getLastRowNum(); i++)
			{
				Row row = sheet.getRow(i);
				if (row == null) {continue;}

				Cell cell = row.getCell(columnIndex);
				columnData.add(getCellValueAsObject(cell));  // 添加单元格值（Object 类型）
			}

			return columnData;
		}
	}

	/**
	 * 统计 Excel 指定列中每个值的出现次数（重复数据频率统计）
	 * 结果按出现次数从高到低排序
	 * @param filePath Excel 文件路径
	 * @param sheetName 工作表名称（支持 null/空字符串 → 默认第一个 Sheet）
	 * @param columnName 列名（表头名称，忽略大小写匹配）
	 * @return List<Map<String, Object>> 格式示例：
	 * [
	 * {"value": "江阴", "count": 45},
	 * {"value": "无锡", "count": 23},
	 * ...
	 * ]
	 */
	public List<Map<String, Object>> countColumnValueFrequency(
			String filePath,
			String sheetName,
			String columnName) throws IOException
	{

		// 复用之前您已有的 readColumnData 方法
		List<Object> columnData = readColumnData(filePath, sheetName, columnName);

		// 使用 Stream 统计频次
		Map<String, Long> frequencyMap = columnData.stream()
												   .map(obj -> obj == null ? "" : obj.toString().trim())  // 统一转字符串 + 去空格
												   .filter(str -> !str.isEmpty())                         // 过滤空值（可选）
												   .collect(Collectors.groupingBy(
														   str -> str,
														   Collectors.counting()
												   ));

		// 转为 List 并按次数降序排序
		return frequencyMap.entrySet().stream()
						   .sorted(Map.Entry.<String, Long>comparingByValue().reversed())
						   .map(entry -> {
							   Map<String, Object> map = new LinkedHashMap<>();
							   map.put("value", entry.getKey());
							   map.put("count", entry.getValue());
							   return map;
						   })
						   .collect(Collectors.toList());
	}

	// 辅助方法：获取单元格值作为 String
	private String getCellValueAsString(Cell cell)
	{
		if (cell == null) {return "";}
		return switch (cell.getCellType())
		{
			case STRING -> cell.getStringCellValue();
			case NUMERIC -> String.valueOf(cell.getNumericCellValue());
			case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
			case FORMULA -> cell.getCellFormula();
			default -> "";
		};
	}

	// 辅助方法：获取单元格值作为 Object（用于 Map）
	private Object getCellValueAsObject(Cell cell)
	{
		if (cell == null) {return null;}
		return switch (cell.getCellType())
		{
			case STRING -> cell.getStringCellValue();
			case NUMERIC -> cell.getNumericCellValue();
			case BOOLEAN -> cell.getBooleanCellValue();
			case FORMULA -> cell.getCellFormula();
			default -> null;
		};
	}
}