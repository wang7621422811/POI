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

