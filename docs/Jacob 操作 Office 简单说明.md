Jacob 使用简介

当前文档转换服务的实现是通过使用 Java 组件 JACOB 调用微软的 office 组件的 COM 接口来实现。

环境准备：
本机安装 Office 软件、jdk、maven。

使用步骤：
1. 从官网上下载 Jacob。
  Jacob 项目地址：https://github.com/freemansoft/jacob-project

  Jacob jar 包下载地址：https://github.com/freemansoft/jacob-project/releases

2. 将下载好的 jacob.jar 文件放入到本地 maven 仓库。
  该 jar 包在 maven 中央仓库中不存在，需手动打到自己的 maven 仓库。
  进入 jar 包所在目录执行以下命令：

  ```
  mvn install:install-file -DgroupId=com.jacob -DartifactId=jacob -Dversion=1.20 -Dfile=jacob.jar -Dpackaging=jar
  ```

3. 把 .dll 文件放入到 `java/jre/bin` 目录下。
  其 .dll 文件是下载 jacob jar 包时附带的文件。

4. 在 maven 项目的 pom 文件中引入 Jacob 依赖。

   ```
   <dependency>
   	<groupId>com.jacob</groupId>
   	<artifactId>jacob</artifactId>
   	<version>1.20</version>
   </dependency>
   ```



Jacob 常用类计方法：

- ComThread：com 组件管理，用于初始化 com 线程，释放线程，可在操作 office 前后使用。
- ActiveXComponent：用于创建 office 应用。
- Dispatch：调度处理类，封装了一些操作来操作 office，里面的所有可操作对象基本都是该类型，Jacob 是一种链式操作模式，
  类似于 StringBuilder 对象，调用 append() 方法返回的还是 StringBuilder 对象。
- Variant：封装参数数据类型。
- Dispatch 的几种静态方法：
    - .call() 方法：调用 com 对象方法，返回 Variant 类型值。
    - .invoke() 方法：和 call() 方法作用相同，但无返回值。
    - .get() 方法：获取 com 对象属性，返回 Variant 类型值。
    - .put() 方法，设置 com 对象属性。
    - 以上方法都有很多重载方法，调用不同的方法需要设置不同的参数，至于哪些参数代表什么意思，具体放什么值，
      就需要参考 office vba 代码，仅靠 Jacob 代码是不行的。
- Variant 对象的 toDispatch() 方法：将上面方法返回的 Variant 类型转换为 Dispatch，进行下一次链式调用。

> 微软 Office VBA API 文档：https://docs.microsoft.com/zh-cn/office/vba/api/overview/



简单的 Word 转 PDF 案例如下：

```
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

public class Demo {
  public static void main(String[] args) {
    String source = "D:\\data\\test.docx"; // 原文件
    String target = "D:\\data\\test.pdf"; // 要转成的文件
    
    // 初始化 com 线程
    // ComThread.InitSTA();
    
    // 创建 Word 应用程序对象
    ActiveXComponent activeXComponent = new ActiveXComponent("Word.application");
    // 设置应用操作文档在后台静默处理，不在明面上显示
    activeXComponent.setProperty("Visible", false);
    // 获取 Documents 对象
    Dispatch documents = activeXComponent.getProperty("Documents").toDispatch();
    // 调用 Documents.Open 方法打开指定的文档
    Dispatch document = Dispatch.call(documents, "Open", source).toDispatch();
    // 调用 Document.SaveAs 方法将文档另存为指定格式
    Dispatch.call(document, "SaveAs", 
      target, // 另存为的文档名称
      17 // 文档的保存格式，17 表示另存为 PDF 格式
    );
    // 调用 Document.Close 方法关闭文档
    Dispatch.call(document, "Close");
    activeXComponent.invoke("Quit");
        
    // 关闭 com 线程
    // ComThread.Release();
  }
}
```

