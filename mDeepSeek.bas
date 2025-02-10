Option Explicit

' ========================================================
' DeepSeek API Integration for Excel
' Author: Sven Bosau
' Website: https://pythonandvba.com
' YouTube: https://youtube.com/@CodingIsFun
' Version: 1.0
' Description: This module provides a custom Excel function
'              to interact with the DeepSeek API for chat completions.
' Requirements: JSON Converter Module (Required for parsing JSON responses)
'               Import the JsonConverter.bas module into your VBA project.
'               Download from: https://github.com/VBA-tools/VBA-JSON
' ========================================================

' API 配置
Private Const API_KEY As String = "YOUR_API_KEY" ' 将其替换为你的 DeepSeek API 密钥
Private Const API_URL As String = "https://api.deepseek.com/v1/chat/completions" ' DeepSeek API 网址
Private Const DEFAULT_MODEL As String = "deepseek-chat" ' AI模型
Private Const DEFAULT_TEMPERATURE As Double = 0.7 ' Temperature创造力的温度，越大越开放 (0 to 2)
Private Const DEFAULT_MAX_TOKENS As Long = 300 ' 最大生成token数( 1 to 8192)
Private Const DEFAULT_TOP_P As Double = 1 ' top_p生成文本时的词汇选择范围 (0 to 1)
Private Const DEFAULT_FREQUENCY_PENALTY As Double = 0 ' 用于惩罚重复的标记 (-2 to 2)
Private Const DEFAULT_PRESENCE_PENALTY As Double = 0 ' 用于用于鼓励新的主题。(-2 to 2)

' HTTP 对象配置
Private Const USE_SERVER_XMLHTTP As Boolean = False
' 设置为真则用 MSXML2.ServerXMLHTTP (用于服务器环境)

' 可用模型model:
' - deepseek-chat: 通用聊天模型
' - deepseek-reasoner: 针对推理任务优化的模型

' 获取 API Key:
' 访问  https://platform.deepseek.com 进行注册从而获得 API Key.

' ========================================================
' 自定义 Excel 函数: DS_Chat
' 描述：向 DeepSeek API 发送一个提示并返回响应。
' Parameters:
' 参数：
' prompt”：这是输入文本或提示，用于提供给 API。
' model”：可选参数，指定要使用的模型，默认值是 “deepseek-chat”。
' temperature”：可选参数，控制随机性，值为 0 时是确定性的，值为 2 时具有创造性。
' max_tokens”：可选参数，指定生成的最大标记数，默认值是 4096。
' top_p”：可选参数，是核采样参数，取值范围是 0 到 1，默认值是 1。
' frequency_penalty”：可选参数，用于惩罚重复的标记，正数更倾向于生成新的增加生成文本的多样性，负数则使模型更倾向于重复。 -2 到 2，默认值是 0。
' presence_penalty”：可选参数，用于鼓励新的主题，正数避免过度重复；负数更频繁地生成某些词汇或内容，-2 到 2，默认值是 0。
' ========================================================

' ========================================================
' 使用示例:
' 以下是在 Excel 中如何使用 DS_Chat 函数的示例：
' 基本用法：

'     通过 “=DS_Chat ("What is the capital of France?")” 调用函数，
'     输出结果: "The capital of France is Paris."

' 高级用法（带自定义参数）：
'     “=DS_Chat ("Explain quantum computing in simple terms.", "deepseek-chat", 0.8, 512, 0.9, 0.5, 0.5)”
'     输入问题并附带一些参数。
'     输出结果："Quantum computing is a type of computing that uses quantum bits (qubits) to perform calculations..."

' 使用定义名称获取 API 密钥：
'     在工作簿中创建名为 “DS_API_KEY” 的定义名称，并在相应单元格中输入 API 密钥。
'     如果命名范围缺失或为空，函数将回退使用常量 API_KEY。

' 注意事项：
'     确保在 VBA 项目引用中启用 “Microsoft XML, v6.0” 和 “Microsoft Scripting Runtime” 库；
'     导入 JsonConverter.bas 模块用于 JSON 解析。

Function DS_Chat(prompt As String, Optional model As String = DEFAULT_MODEL, _
                 Optional temperature As Double = DEFAULT_TEMPERATURE, Optional max_tokens As Long = DEFAULT_MAX_TOKENS, _
                 Optional top_p As Double = DEFAULT_TOP_P, Optional frequency_penalty As Double = DEFAULT_FREQUENCY_PENALTY, _
                 Optional presence_penalty As Double = DEFAULT_PRESENCE_PENALTY) As String
    On Error GoTo ErrorHandler
    
    ' Validate input 验证输入
    If prompt = "" Then
        DS_Chat = "Error: Prompt cannot be empty."
        Exit Function
    End If
    
    ' Optimize performance 优化性能
    Application.ScreenUpdating = False ' 屏幕刷新
    Application.DisplayStatusBar = True ' 状态栏显示
    Application.StatusBar = "Calling DeepSeek API..." ' 状态栏信息
    
    ' Create HTTP object based on configuration 根据配置创建 HTTP 对象
    Dim http As Object
    If USE_SERVER_XMLHTTP Then
        Set http = CreateObject("MSXML2.ServerXMLHTTP") ' Use this for server environments 用于服务器环境
    Else
        Set http = CreateObject("MSXML2.XMLHTTP") ' Use this for desktop environments 用于桌面环境
    End If
    
    ' Get API key from named range or use the constant 从定义名称获取 API 密钥或使用常量。
    Dim apiKey As String
    On Error Resume Next
    apiKey = ActiveWorkbook.Names("DS_API_KEY").RefersToRange.Value
    On Error GoTo ErrorHandler
    If apiKey = "" Then
        apiKey = API_KEY ' Fallback to the constant if named range is missing or empty
    End If
    
    ' Prepare the request body 准备请求体
    Dim requestBody As String
    requestBody = "{""model"": """ & model & """, ""messages"": [{""role"": ""user"", ""content"": """ & prompt & """}], " & _
                  """temperature"": " & temperature & ", ""max_tokens"": " & max_tokens & ", ""top_p"": " & top_p & ", " & _
                  """frequency_penalty"": " & frequency_penalty & ", ""presence_penalty"": " & presence_penalty & "}"
    
    ' Send the request 发送请求
    http.Open "POST", API_URL, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send requestBody
    
    ' Check the response status 检查相应状态
    If http.Status = 200 Then
        ' Parse the response 解析响应
        Dim response As String
        response = http.responseText
        Dim json As Object
        Set json = JsonConverter.ParseJson(response)
        
        ' Extract the completion text 提取完成文本
        DS_Chat = json("choices")(1)("message")("content")
    Else
        ' Handle API errors 处理API错误
        DS_Chat = "API Error: " & http.Status & " - " & http.statusText
    End If
    
    ' Clear status bar and restore settings 清除状态栏并回复设置
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Exit Function
    
ErrorHandler:
    ' Handle VBA errors 处理VBA错误
    DS_Chat = "VBA Error: " & Err.Description
    ' Clear status bar and restore settings 清除状态栏并回复设置
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Function
