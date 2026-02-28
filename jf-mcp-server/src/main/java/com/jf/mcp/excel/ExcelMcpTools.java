package com.jf.mcp.excel;

import lombok.extern.slf4j.Slf4j;
import org.springframework.ai.tool.ToolCallbackProvider;
import org.springframework.ai.tool.annotation.Tool;
import org.springframework.ai.tool.annotation.ToolParam;
import org.springframework.ai.tool.method.MethodToolCallbackProvider;
import org.springframework.context.annotation.Bean;
import org.springframework.stereotype.Component;

import java.util.Collections;
import java.util.List;
import java.util.Map;

@Component
@Slf4j
public class ExcelMcpTools
{
	private final ApachePoiExcelReader excelReader;

	public ExcelMcpTools(ApachePoiExcelReader excelReader)
	{
		this.excelReader = excelReader;
	}

	@Bean
	public ToolCallbackProvider excelTools(ExcelMcpTools excelMcpTools)
	{
		return MethodToolCallbackProvider.builder().toolObjects(excelMcpTools).build();
	}

	@Tool(description = "读取Excel指定工作表的所有数据，返回结构化JSON")
	public String readExcelSheet(
			@ToolParam(description = "Excel 文件的本地路径（字符串，必填，例如 /path/to/file.xlsx）") String filePath,
			@ToolParam(description = "工作表名称（字符串，必填，例如 Sheet1；如果为空，可默认第一个工作表）") String sheetName,
			@ToolParam(description = "表头行号（整数，必填，0-based，例如 0 表示第一行）") int headerRow)
	{
		try
		{
			log.info("ExcelMcpTools-读取Excel所有数据,filePath:{},sheetName:{},headerRow:{}", filePath, sheetName, headerRow);
			String content = excelReader.readSheetAsJson(filePath, sheetName, headerRow);
			log.info("ExcelMcpTools-读取Excel所有数据:{}", content);
			return content;
		}
		catch (Exception e)
		{
			log.error("ExcelMcpTools-读取Excel所有数据失败,e:", e);
			return "Error: " + e.getMessage();
		}
	}

	@Tool(description = "精确统计Excel中某个关键词/数值出现的次数（支持正则）")
	public int countInExcel(
			@ToolParam(description = "Excel 文件的本地路径（字符串，必填，例如 /path/to/file.xlsx）") String filePath,
			@ToolParam(description = "工作表名称（字符串，必填，例如 Sheet1；如果为空，可默认第一个工作表）") String sheetName,
			@ToolParam(description = "要统计的关键词或数值（字符串，必填）") String keyword,
			@ToolParam(description = "是否使用正则表达式匹配（布尔值，必填，true 为使用，false 为简单包含匹配）") boolean useRegex)
	{
		try
		{
			log.info("ExcelMcpTools-统计Excel中关键词数量,filePath:{},sheetName:{},keyword:{},useRegex:{}", filePath, sheetName, keyword, useRegex);
			int i = excelReader.countOccurrences(filePath, sheetName, keyword, useRegex);
			log.info("ExcelMcpTools-统计Excel中关键词数量:{}", i);
			return i;
		}
		catch (Exception e)
		{
			log.error("ExcelMcpTools-统计Excel中关键词数量出错,e:", e);
			return -1;  // 错误时返回 -1
		}
	}

	@Tool(description = "按条件过滤Excel行，返回匹配数据")
	public List<Map<String, Object>> filterExcelRows(
			@ToolParam(description = "Excel 文件的本地路径（字符串，必填，例如 /path/to/file.xlsx）") String filePath,
			@ToolParam(description = "工作表名称（字符串，必填，例如 Sheet1；如果为空，可默认第一个工作表）") String sheetName,
			@ToolParam(description = "要过滤的列名（字符串，必填，假设第 0 行为表头）") String column,
			@ToolParam(description = "匹配的值（字符串，必填，忽略大小写匹配）") String value)
	{
		try
		{
			log.info("ExcelMcpTools-按条件过滤Excel行,filePath:{},sheetName:{},column:{},value:{}", filePath, sheetName, column, value);
			List<Map<String, Object>> maps = excelReader.filterRows(filePath, sheetName, column, value);
			log.info("ExcelMcpTools-按条件过滤Excel行,查询到数据:{}", maps);
			return maps;
		}
		catch (Exception e)
		{
			log.error("ExcelMcpTools-按条件过滤Excel行出错,e:", e);
			return Collections.emptyList();  // 错误时返回空列表
		}
	}
	@Tool(description = "统计Excel指定列中每个值的出现次数（建议先查询Excel元数据）,结果按出现次数从高到低排序")
	public List<Map<String, Object>> countColumnValueFrequency(
			@ToolParam(description = "Excel 文件的本地路径（字符串，必填，例如 /path/to/file.xlsx）") String filePath,
			@ToolParam(description = "工作表名称（字符串，必填，例如 Sheet1；如果为空，可默认第一个工作表）") String sheetName,
			@ToolParam(description = "要查询的列名（字符串，必填，假设第 0 行为表头）") String column
			)
	{
		try
		{
			log.info("ExcelMcpTools-统计Excel指定列中每个值的出现次数,filePath:{},sheetName:{},column:{}", filePath, sheetName, column);
			List<Map<String, Object>> maps = excelReader.countColumnValueFrequency(filePath, sheetName, column);
			log.info("ExcelMcpTools-统计Excel指定列中每个值的出现次数,查询到数据:{}", maps);
			return maps;
		}
		catch (Exception e)
		{
			log.error("ExcelMcpTools-统计Excel指定列中每个值的出现次数,e:", e);
			return Collections.emptyList();  // 错误时返回空列表
		}
	}
	@Tool(description = "读取Excel元数据：工作表列表、列名、行数等(建议在处理前都先调用该方法获取Excel元数据)")
	public String getExcelMetadata(
			@ToolParam(description = "Excel 文件的本地路径（字符串，必填，例如 /path/to/file.xlsx）") String filePath)
	{
		try
		{
			log.info("ExcelMcpTools-读取Excel元数据,filePath:{}", filePath);
			String excelMetadata = excelReader.getExcelMetadata(filePath);
			log.info("ExcelMcpTools-读取Excel元数据,excelMetadata:{}", excelMetadata);
			return excelMetadata;
		}
		catch (Exception e)
		{
			log.error("ExcelMcpTools-读取Excel元数据,e:", e);
			return "Error: " + e.getMessage();
		}
	}
}