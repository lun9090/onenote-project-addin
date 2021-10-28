# onente-all-in-one
onenote项目管理插件
## 背景
本人一直想找一个能对工作项状态进行管理，能够描述文字，又能在文字下面保存文件的笔记，一直用印象笔记很多年，今年用印象笔记不好用，感觉很臃肿，而且很慢，功能限制太厉害，最终选择了OneNote，找了很多插件，目前在用的插件有onemore和OnenoteTaggingKit插件，onemore用的比较多的两个功能是查看xml和样式功能，OnenoteTaggingKit不能满足我的需求，我需要在标题上能看到标签，所以想到了自己开发。

# 功能清单

- 标记管理

  - 更新标题标记

    > 格式：{标记}｜...｜{标题}。

  - 新增标记

    - 【未开展】
    - 【开展中】
    - 【未确认】
    - 【作废】
    - 【待设计】
    - 【未转】
    - 【合并】
    - 【已转】
    - 【暂不开展】
    - 【已转需补充】
    - 【已完成】

  - 删除标记

    - 【未开展】
    - 【开展中】
    - 【未确认】
    - 【作废】
    - 【待设计】
    - 【未转】
    - 【合并】
    - 【已转】
    - 【暂不开展】
    - 【已转需补充】
    - 【已完成】

  - 删除所有标记

- 综合管理

  - 查看xml

    > 查看本页xml，且能够复制。

- 页面管理

  - 删空数据

    > 删除没有内容的内容块

  - 统一数据位置

    > 统一唯一一个含有内容块的大小及位置
    >
    > 需要优化增加多个含有内容块场景

# 特别说明

本工具根据个人工作需要开发，有其他需求的，可以留言。

## 开发环境

- Microsoft Visual Studio Community 2019 版本 16.11.4
  - Microsoft Visual Studio Installer Projects
  - .NET Framework 4.8
- Microsoft® OneNote® 适用于 Microsoft 365 MSO (16.0.14228.20288) 64 位



## 特别感谢

因为他们的付出，才有我今天的成果。

- 创建工程说明：https://www.cnblogs.com/JohnHwangBlog/p/6305380.html

- OneNote2007开发说明：https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/ms788684(v=office.12)?redirectedfrom=MSDN

- OneNote2010开发说明：https://docs.microsoft.com/zh-cn/archive/msdn-magazine/2010/july/onenote-2010-creating-onenote-2010-extensions-with-the-onenote-object-model

- OneNote开发人员参考：https://docs.microsoft.com/zh-cn/office/client-developer/onenote/onenote-home

- C# 截取字符串某个字符分割的最后一部分：https://www.cnblogs.com/wdd812674802/p/10406956.html

- C语言break和continue用法详解（跳出循环）：http://c.biancheng.net/view/1812.html

- xml相关：https://docs.microsoft.com/en-us/dotnet/api/system.xml.linq?view=net-5.0

- XDocument descendants（foreach 的使用）：https://stackoverflow.com/questions/22492902/xdocument-descendants

- 获取xdocument中的节点属性：http://cn.voidcc.com/question/p-kxkrfvnw-bdo.html

- C# 使用XDocument实现读取、添加，修改XML文件：https://www.cnblogs.com/haosit/p/6801420.html

- 从XDocument中删除节点：https://www.thinbug.com/q/3215470

- c# – 向Xdocument添加新的XElement：http://www.voidcn.com/article/p-bnooxczw-bvm.html

- 图标来源：https://iconstore.co/

- C# 如何获取时间各种方法（日期+具体时间）：https://www.cnblogs.com/qy1234/p/12170612.html

- c#调用类方法时，被引用的类 有无public修饰问题：https://zhidao.baidu.com/question/1739128861328867907.html

- C#字符串转换为数字的4种方法：https://blog.csdn.net/coolszy/article/details/83531866

- C#中字符串与数值的相互转换：https://www.cnblogs.com/hans_gis/archive/2011/04/16/2018318.html

- 图标相关：https://github.com/WetHat/OnenoteTaggingKit/blob/54e7f263d445cc0e7e190c26facc61a0fdfa02f0/OneNoteTaggingKit/Connect.cs

- 修复xml：

  - https://zhidao.baidu.com/question/509973126.html 
  - ed2k://|file|cn_msxml_4.0_service_pack_3_x86.msi|2373120|ABFEF286E3620313057B222B1699A732|/

- VSTO开发入门，使用CustomUI自定义Office功能区：http://cas01.com/6482.html

- VSTO开发入门，CustomUI元素详解：https://zhuanlan.zhihu.com/p/338524994

- Excel CustomUI功能区布局：http://cn.voidcc.com/question/p-ejdgwjjv-kk.html

- C#比较两个list集合，两集合同时存在或A集合存在B集合中无：https://blog.csdn.net/smartsmile2012/article/details/54408439/

- C#创建Windows窗体应用程序（WinForm程序）：http://c.biancheng.net/view/2945.html

- 在应用程序中创建第一个 IWin32Window 对象之前，必须调用 SetCompatibleTextRenderingDefault。...：https://blog.csdn.net/ants717007/article/details/101120199

- 周末浅说--未将对象引用设置到对象的实例(System.NullReferenceException)：https://www.cnblogs.com/cyq1162/archive/2011/07/24/2115388.html

- C#中判断字符串为空的几种方法的比较：https://blog.csdn.net/biaobiao1217/article/details/39047963

- 使用cmd启动exe文件：https://blog.csdn.net/wl724120268/article/details/84846884

- C#中string转成int类型：https://blog.csdn.net/shengyingpo/article/details/76618681

- 如何把字符串转换成数字(带小数)：https://bbs.csdn.net/topics/290040564?list=977565

- C# winform 用textbox显示文本 如何把光标定位到指定的位置：https://www.cnblogs.com/winformasp/articles/11903572.html

- C# 教程：https://www.runoob.com/csharp/csharp-tutorial.html

- C#各种异常处理方式：https://www.cnblogs.com/darrenji/p/3965443.html

- 如何在C#中让一个过程等待1秒钟后再执行下面的语句?：https://bbs.csdn.net/topics/80202282

- C# 对象集合初始化：https://www.cnblogs.com/lgxlsm/p/10950135.html

- C# 循环break 和continue：https://www.cnblogs.com/winward996/p/11502481.html

- 如何使用XDocument删除节点和子节点(How to delete nodes and subnodes using XDocument)：https://www.it1352.com/1558565.html

  



