# onenote-project-addin
onenote项目管理插件
## 背景
本人一直想找一个能对工作项状态进行管理，能够描述文字，又能在文字下面保存文件的笔记，一直用印象笔记很多年，今年用印象笔记不好用，感觉很臃肿，而且很慢，功能限制太厉害，最终选择了OneNote，找了很多插件，目前在用的插件有onemore和OnenoteTaggingKit插件，onemore用的比较多的两个功能是查看xml和样式功能，OnenoteTaggingKit不能满足我的需求，我需要在标题上能看到标签，所以想到了自己开发。
随着onenote使用的不断深入，新增了不少功能，优化了不少功能，见功能清单。

## 功能清单
![image](https://user-images.githubusercontent.com/78783303/215311541-862f0843-9d39-4678-bad0-c0b6676ae5f2.png)


## 需要解释的功能

- 标记管理

  - 更新标题标记

    > 格式：{标记}｜...｜{标题}。
    > 标记管理：插件可以配合onenote的标记功能使用，可以设置标记，设置后，点击更新标题标记即可


- 综合管理

  - 查看xml

    > 查看本页xml，且能够复制。

- 页面管理

  - 删空数据

    > 删除没有内容的内容块

  - 统一数据位置（A4宽）

    > 统一含有内容块的大小及位置
    > 删除空的内容块
    > 如果存在标签（onemore中的页面标签和OnenoteTaggingKit插件生成的标签），跳过标签内容块。
    > 设置内容块的宽度为451.2755737304687，来适配打印时的边距。
  - A4页面设置
    > 点击后，当前页面大小设置为A4，慎用，限制住了内容数量，可以使用统一数据位置（A4宽）

- 日记管理
    - 创建日记页（MyJournal.Notebook）
        > 很粗糙，点击后生成今日的日志页。
    - 项目管理日记
        > 点击后，在默认笔记本位置，生成My Project Journal笔记本。
        > 如果的笔记本是联网的，也就是说，你的默认笔记本路径是在onedrve上的，按照这样的步骤操作，第一次生成My Project Journal笔记本后，关闭OneNote，等待onedirve将你的笔记本自动转化为在线笔记本，这个时间可能有点慢，你也可以继续使用，等待下次打开找不到笔记本的时候，你需要做的一件事就是，点击打开笔记本，点击onedirve，找到转化后的在线My Project Journal笔记本，完成同步后，插件可继续使用，目前技术还未解决该问题，后续再说。
- 右键功能
    - 复制为纯文本
        > 选择内容，右键，点击复制为纯文本，将文本复制到剪切板。
## 特别说明

本工具根据个人工作需要开发，有其他需求的，可以留言。

## 开发环境

- Microsoft Visual Studio Community 2022 (64 位) 
  - Microsoft Visual Studio Installer Projects
  - .NET Framework 4.8
- Microsoft® OneNote® 适用于 Microsoft 365 MSO (16.0.14228.20288) 64 位
