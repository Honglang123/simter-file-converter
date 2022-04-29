# 文件转换 Rest API

通过属性 `simter-file-converter.rest-context-path` 配置事故 Rest 路径的根路径，默认值为 `/file-converter`。以下的 URL 均使用此默认值。

## 1. 执行文件转换

**请求：**

```
POST /file-converter?from-file=x&to-file=x&password=x
```

| Name      | Require | Description
|-----------|---------|-------------
| from-file | true    | 来源文件的相对路径
| to-file   | true    | 转换后文件的相对路径
| password  | false   | 来源文件设置的保护密码

**响应（成功）：**

```
204 No Content
```

如果来源文件不存在，响应返回：

```
410 Gone
Content-Type : plain/text

来源文件不存在！
```