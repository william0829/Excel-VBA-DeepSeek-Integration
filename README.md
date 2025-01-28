
# DeepSeek API Integration for Excel

Use this Excel module to send chat prompts to the DeepSeek API. You can customize the prompt, model, and other parameters right in your worksheet.

## Features
- Send chat prompts to DeepSeek with a simple Excel function.
- Use optional parameters for more control over responses.
- Store your DeepSeek API key in a named range (`DS_API_KEY`) or in the code.

## How It Works
1. Import the `JsonConverter.bas` file into your VBA project.  
2. Import the `mDeepSeek.bas` file into your VBA project.  
3. Add your API key in `mDeepSeek.bas`:
   ```vba
   Private Const API_KEY As String = "YOUR_API_KEY"
   ```
   or create a named range in Excel called `DS_API_KEY` and place your key there.
4. Call the function `DS_Chat` in any cell:
   ```excel
   =DS_Chat("Hello DeepSeek!")
   ```
   This returns the response from the API.

## Optional Parameters
Use these optional parameters to fine-tune your prompt:
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
- **model:** Choose `"deepseek-chat"` or `"deepseek-reasoner"`.
- **temperature:** Adjust creativity (0–2).
- **max_tokens:** Set token limit.
- **top_p:** Use nucleus sampling (0–1).
- **frequency_penalty:** Penalize repeated tokens (–2 to 2).
- **presence_penalty:** Encourage new topics (–2 to 2).

## 🤝 Connect with Me
- 📺 **YouTube:** [CodingIsFun](https://youtube.com/c/CodingIsFun)
- 🌐 **Website:** [PythonAndVBA](https://pythonandvba.com)
- 💬 **Discord:** [Join the Community](https://pythonandvba.com/discord)
- 💼 **LinkedIn:** [Sven Bosau](https://www.linkedin.com/in/sven-bosau/)
- 📸 **Instagram:** [sven_bosau](https://www.instagram.com/sven_bosau/)

## 💖 Support
If my tutorials help you, please consider [buying me a coffee](https://pythonandvba.com/coffee-donation).  
[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://pythonandvba.com/coffee-donation)

## 📬 Feedback & Collaboration
If you have ideas, feedback, or want to collaborate, reach out at contact@pythonandvba.com.  
![Logo](https://www.pythonandvba.com/banner-img)
