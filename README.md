非原创！
# DeepSeek API Integration for Excel 🐋

使用此 Excel 模块将聊天提示发送到 DeepSeek API。您可以在工作表中直接自定义提示、模型和其他参数。


## Video Tutorial
[![YouTube Video](https://img.youtube.com/vi/ln8oxm9Gvjs/0.jpg)](https://youtu.be/ln8oxm9Gvjs)

## 功能
- 通过一个简单的 Excel 函数将聊天提示发送到 DeepSeek。
- 使用可选参数以更好地控制响应。
- 将你的 DeepSeek API 密钥存储在名称（`DS_API_KEY`）或代码中。

## 工作原理
1. 将`JsonConverter.bas` 文件导入到你的 VBA 项目中。
2. 将 `mDeepSeek.bas` 文件导入到你的 VBA 项目中。
3. 在 `mDeepSeek.bas`中添加你的 API 密钥：
   ```vba
   Private Const API_KEY As String = "YOUR_API_KEY"
   ```
   或者在 Excel 中创建一个名为 `DS_API_KEY` 的名称，并将你的密钥放在那里。
4. 在任何单元格中调用函数 `DS_Chat` ：
   ```excel
   =DS_Chat("你好，DeepSeek！")
   ```
   它将返回来自 API 的响应。


## 可选参数
使用这些可选参数来微调你的提示：
```excel
=DS_Chat(
   prompt, 
   [model], 
   [temperature], 
   [max_tokens], 
   [top_p], 
   [frequency_penalty], 
   [presence_penalty]
)
```
- **model:** 选择 `"deepseek-chat"` 或 `"deepseek-reasoner"`.
- **temperature:** 调整创造力 (0–2).
- **max_tokens:** 设置 token 限制.
- **top_p:** Use nucleus sampling (0–1).
- **frequency_penalty:** 惩罚重复的标记。 (–2 to 2).
- **presence_penalty:** 鼓励新主题 (–2 to 2).

## 🤝 原作者
- 📺 **YouTube:** [CodingIsFun](https://youtube.com/c/CodingIsFun)
- 🌐 **Website:** [PythonAndVBA](https://pythonandvba.com)
- 💬 **Discord:** [Join the Community](https://pythonandvba.com/discord)
- 💼 **LinkedIn:** [Sven Bosau](https://www.linkedin.com/in/sven-bosau/)
- 📸 **Instagram:** [sven_bosau](https://www.instagram.com/sven_bosau/)
