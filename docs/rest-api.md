# 文件类型转换 Rest API

通过属性 `simter-file-converter.rest-context-path` 配置事故 Rest 路径的根路径，默认值为 `/file-converter`。以下的 URL 均使用此默认值。

## 1. 将其他类型文件转换为 pdf 文件并返回。

**请求：**

```
请求行：
POST /convert2pdf/?filename=x

请求头：
password: $password
Content-Type: $contentType

请求体：
$fileData
```

请求参数说明：
| Name        | Require | Description
|-------------|---------|-------------
| filename    | false   | 源文件名（不含文件扩展名）
| password    | false   | 源文件加密密码
| contentType | true    | 源文件对应的实际 MIME 类型。
| fileData    | true    | 源文件二进制数据

office 文件所对应的 MIME 类型如下：
| 文件后缀 | MIME TYPE
|--------|---------------
| .doc   | application/msword
| .dot	 | application/msword
| .docx	 | application/vnd.openxmlformats-officedocument.wordprocessingml.document
| .dotx	 | application/vnd.openxmlformats-officedocument.wordprocessingml.template
| .docm	 | application/vnd.ms-word.document.macroEnabled.12
| .dotm	 | application/vnd.ms-word.template.macroEnabled.12
| .xls	 | application/vnd.ms-excel
| .xlt	 | application/vnd.ms-excel
| .xla	 | application/vnd.ms-excel
| .xlsx	 | application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
| .xltx	 | application/vnd.openxmlformats-officedocument.spreadsheetml.template
| .xlsm	 | application/vnd.ms-excel.sheet.macroEnabled.12
| .xltm	 | application/vnd.ms-excel.template.macroEnabled.12
| .xlam	 | application/vnd.ms-excel.addin.macroEnabled.12
| .xlsb	 | application/vnd.ms-excel.sheet.binary.macroEnabled.12
| .ppt	 | application/vnd.ms-powerpoint
| .pot	 | application/vnd.ms-powerpoint
| .pps	 | application/vnd.ms-powerpoint
| .ppa	 | application/vnd.ms-powerpoint
| .pptx	 | application/vnd.openxmlformats-officedocument.presentationml.presentation
| .potx	 | application/vnd.openxmlformats-officedocument.presentationml.template
| .ppsx	 | application/vnd.openxmlformats-officedocument.presentationml.slideshow
| .ppam	 | application/vnd.ms-powerpoint.addin.macroEnabled.12
| .pptm	 | application/vnd.ms-powerpoint.presentation.macroEnabled.12
| .potm	 | application/vnd.ms-powerpoint.presentation.macroEnabled.12
| .ppsm	 | application/vnd.ms-powerpoint.slideshow.macroEnabled.12

**响应（成功）：**

```
200 OK
Content-Type        : application/pdf
Content-Disposition : attachment; filename="$filename"

$convertedData
```

>注：Content-Disposition 值中参数 filename 的值必需用双引号引住，并且必须使用 ISO-8859-1 字符编码（RFC2183）。