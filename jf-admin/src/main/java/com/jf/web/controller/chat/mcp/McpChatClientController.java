package com.jf.web.controller.chat.mcp;

import com.jf.common.core.domain.AjaxResult;
import org.apache.commons.io.FileUtils;
import org.springframework.ai.chat.client.ChatClient;
import org.springframework.ai.tool.ToolCallback;
import org.springframework.ai.tool.ToolCallbackProvider;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;

/**
 * @author CYF
 * @description MCP工具调用
 * @Date 2026/2/27
 */
@RestController
@RequestMapping("/api/mcp")
public class McpChatClientController
{
	private static final String CHAT_FILE_PATH = "/data/jf/chat/file/";

	private final ChatClient chatClient;

	private final ToolCallbackProvider toolCallbackProvider;

	public McpChatClientController(ChatClient.Builder builder, ToolCallbackProvider toolCallbackProvider)
	{
		chatClient = builder.build();
		this.toolCallbackProvider = toolCallbackProvider;
	}

	@PostMapping("/chat/excel")
	public AjaxResult chatByExcel(@RequestParam("file") MultipartFile file, @RequestParam("text") String text)
	{
		if (file.isEmpty())
		{
			return AjaxResult.error("请选择文件");
		}
		String filename = file.getOriginalFilename() == null ? file.getName() : file.getOriginalFilename();
		String filepath = CHAT_FILE_PATH + filename;
		String content;
		try
		{
			File tempFile = new File(filepath);
			FileUtils.copyInputStreamToFile(file.getInputStream(), tempFile);
			text = text + "\n 文件路径:" + filepath;
			ToolCallback[] toolCallbacks = toolCallbackProvider.getToolCallbacks();
			content = chatClient.prompt(text)
								.toolCallbacks(toolCallbacks)
								.call().content();
		}
		catch (IOException e)
		{
			throw new RuntimeException(e);
		}
		return AjaxResult.success(content);
	}
}
