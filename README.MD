# 5G手机评价分析软件

基于**爬虫**与**AI技术**的京东平台5G手机评论自动化分析系统



****

## 项目概述

5G手机评价分析软件是一款专为京东平台5G手机评论深度挖掘而设计的自动化分析工具，通过用户输入的商品ID抓取京东商品评论API中的评论数据，并通过去重、清洗、分词等预处理步骤提高数据质量。在自然语言处理方面，系统利用jieba库进行中文分词，snownlp库进行情感倾向分析，量化商品评价的正负情感比例。同时，系统接入了通义千问大语言模型API，对处理后的评论数据进行深层次分析，生成详尽的京东平台5G手机分析报告。

### 项目结构

```plaintext
5G手机评价/
        

└── main.py  

├── data
│   ├── bg.png        # 词云图背景图片
│   └── stopwords.txt # 停用词表

└── output
    ├── cleaned_data.xlsx  # 清理后整理的评论数据集
    ├── contents.txt      # 摘要和总结的评论内容
    ├── original_data.xlsx # 爬取得到的原始评论数据
    ├── report.docx       # 自动生成的综合分析报告
    ├── segmented_words.txt # 分词后的词汇列表
    ├── sentiment_analysis_pie_chart.png # 情感分析结果饼状图
    └── wordcloud.png        # 关键词词云图
```

### 系统需求
- **Python环境**:Python 3.11或更新版本

### 安装指南

#### Pandoc安装

为了生成Word文档报告，请先安装Pandoc，这是一个强大的文档格式转换工具。访问[Pandoc官网安装指南](https://pandoc.org/installing.html)以获取适合您操作系统的安装方法。

### 使用步骤

1. **克隆项目**:从GitHub或其他托管平台克隆本项目至本地计算机，然后使用PyCharm或其他兼容IDE打开该项目。

2. **安装依赖**:在项目根目录下执行以下命令安装所有必需的第三方库:
   ```shell
   pip install -r requirements.txt
   ```

3. **配置API密钥**:在`qianwen.py`文件中填写您的通义千问API密钥:
   ```python
   dashscope.api_key = "在此处输入你的API密钥"
   ```
4. **运行分析**:启动主程序`main.py`，根据提示输入待分析的京东商品链接，点击**开始分析**按钮，等待程序执行完毕。

![alt text](image.png)

![alt text](image-1.png)

5. **查看结果**:分析完成后，进入f`{product_id}`目录，可找到包含如下内容的各类输出文件:

   - `cleaned_data.xlsx`:已清洗和整理过的评论数据
   - `contents.txt`:提炼总结的评论主要内容
   - `original_data.xlsx`:爬取得到的原始评论记录
   - `report.docx`:系统自动生成的分析报告（利用Pandoc转换生成）
   - `segmented_words.txt`:经过分词处理的词语列表
   - `sentiment_analysis_pie_chart.png`:情感倾向分析的饼状统计图
   - `wordcloud.png`:反映高频关键词的词云图像
   - 
****
### 本项目遵循MIT许可协议开源。