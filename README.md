# 1.POI报表的概述
在企业级应用开发中，Excel报表是一种最常见的报表需求。Excel报表开发一般分为两种形式：
- 为了方便操作，基于Excel的报表批量上传数据
- 通过java代码生成Excel报表

## 1.1 Excel报表的两种形式

 |      |Excel2003|Excel2007|
 |------|--------|---------|
 |后缀|xls|xlsx|
 |结构|二进制结构,其核心结构是符合性结构|XMl类型结构|
 |单sheet|行:65535;列:256|行:1048576;列:16384|
 |特点|存储容量有限|基于xml压缩占用空间小|

## 1.2 常见的Excel操作工具
-> Java中常见的用来操作Excl的方式一般有2种：JXL和POI
- JXL只能对Excel进行操作,属于比较老的框架，它只支持到Excel 95-2000的版本。现在已经停止更新和维护。
- POI是apache的项目,可对微软的Word,Excel,Ppt进行操作,包括office2003和2007,Excl2003和2007。poi现在
一直有更新。所以现在主流使用POI。

## 1.3 POI的概述
Apache POI是Apache软件基金会的开源项目，由Java编写的免费开源的跨平台的 Java API，Apache POI提供API
给Java语言操作Microsoft Office的功能

## 1.4 POI的应用场景
1. 数据报表生成
2. 数据备份
3. 数据批量上传

# 2 POI的入门操作

## 2.1 搭建环境

```xml
<dependencies>
  <dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>4.0.1</version>
  </dependency>
  <dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>4.0.1</version>
  </dependency>
  <dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml-schemas</artifactId>
    <version>4.0.1</version>
  </dependency>
</dependencies>
```

## 2.2 POI结构说明

- **HSSF**提供读写Microsoft Excel XLS格式档案的功能。
- **XSSF**提供读写Microsoft Excel OOXML XLSX格式档案的功能。
- **HWPF**提供读写Microsoft Word DOC格式档案的功能。
- **HSLF**提供读写Microsoft PowerPoint格式档案的功能。
- **HDGF**提供读Microsoft Visio格式档案的功能。
- **HPBF**提供读Microsoft Publisher格式档案的功能。
- **HSMF**提供读Microsoft Outlook格式档案的功能。

## 2.3  API介绍


|API名称    |                                 |
|-----------|---------------------------------|
|Workbook   |Excel的文档对象,针对不同的Excel类型分为：HSSFWorkbook（2003）和XSSFWorkbool（2007）|
|Sheet Excel|的表单      |
|Row Excel的|行          |
|Cell Excel |的格子单元  | 
|Font Excel |字体        |
|CellStyle  |格子单元样式|

#3. 模板打印
自定义生成Excel报表文件还是有很多不尽如意的地方，特别是针对复杂报表头，单元格样式，字体等操作。手写
这些代码不仅费时费力，有时候效果还不太理想。那怎么样才能更方便的对报表样式，报表头进行处理呢？答案是
使用已经准备好的Excel模板，只需要关注模板中的数据即可。


#4.百万级数据内存溢出
[导出百万级数据内存溢出](images/导出百万级数据内存溢出.png)
##4.1解决方案分析
对于百万数据量的Excel导入导出，只讨论基于Excel2007的解决方法。在ApachePoi 官方提供了对操作大数据量的
导入导出的工具和解决办法，操作Excel2007使用XSSF对象，可以分为三种模式：
- 用户模式：用户模式有许多封装好的方法操作简单，但创建太多的对象，非常耗内存（之前使用的方法）
- 事件模式：基于SAX方式解析XML，SAX全称Simple API for XML，它是一个接口，也是一个软件包。它是一
种XML解析的替代方法，不同于DOM解析XML文档时把所有内容一次性加载到内存中的方式，它逐行扫描文
档，一边扫描，一边解析。
- SXSSF对象：是用来生成海量excel数据文件，主要原理是借助临时存储空间生成excel

#5.百万数据报表读取
[POI读取百万数据报表](images/POI读取百万数据报表.png)

##5.1思路分析
- 用户模式：加载并读取Excel时，是通过一次性的将所有数据加载到内存中再去解析每个单元格内容。当Excel
数据量较大时，由于不同的运行环境可能会造成内存不足甚至OOM异常。
- 事件模式：它逐行扫描文档，一边扫描一边解析。由于应用程序只是在读取数据时检查数据，因此不需要将
数据存储在内存中，这对于大型文档的解析是个巨大优势。
#5.2 步骤分析
设置POI的事件模式
根据Excel获取文件流
根据文件流创建OPCPackage
创建XSSFReader对象
（2）Sax解析
自定义Sheet处理器
创建Sax的XmlReader对象
设置Sheet的事件处理器
逐行读取