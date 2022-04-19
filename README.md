# 基于模板的 word 自动生成工具

依赖`python-docx`，实现了基于模板对 word 文档的自动生成。本文定义了一种极为简单的模板语法，无学习成本，对非开发人员也极为友好。

Relying on `python-docx`, it realizes the automatic generation of word documents based on templates. This article defines an extremely simple template syntax that has no learning cost and is extremely friendly to non-developers.



## 使用方法

### 步骤

1. 首先自定义模板，并在待插入数据的地方使用内容标签占位，可参考 data/template-demo.docx
2. 以字典方式提供数据，每种插入类型的数据格式不同
3. 指定模板路径、生成路径与生成文件名，执行并输出异常信息，得到生成的word文件

生成过程如下：

模板内容：

```
我的朋友是{{text:friend}}。
```

传入参数数据：

```
dict = {
    "friend": "小明",
}
```

文档生成结果：

```
我的朋友是小明。
```



### 模板规则

待插入内容采用 `{{type:name}}` 内容标签格式，例如：`{{text:t1}}`。工具会直接使用实际内容替换内容标签，不修改模板内容部分的格式与其他内容，但要求内容标签的格式完全相同，否则不会识别该标签。

- name: 内容标签名，需要模板内全局唯一，不符合条件的模板无法载入

- type: 预定义的内容标签类型，模板会忽略无法识别的类型，如果不指定，则默认为 `text`

  类型主要分为两种：无内容类型与有内容类型。
  - 无内容类型会自动根据当前情况填写信息，`name`自动忽略，可省略`name`，但不能省略`:`
  
    常见的类型有：
    - date: 当前日期，默认格式为 `YYYY-MM-DD`
    - time: 当前时间，默认24小时制，格式为 `HH:MM:SS`
  
  - 有内容类型需要根据 `name` 在生成文档时指定填写的内容。
  
    常见的类型有：
    - text: 文本
    - unordered-list: 无序列表
    - ordered-list: 有序列表
    - image: 图片
    - table: 表格
    - link: 链接



### 文档生成说明

插入内容采用字典格式，key为模板中内容标签的`name`，`value`为实际内容。不同类型的插入内容`value`格式不同。

插入内容字典中模板不需要的内容会被自动忽略。字典中缺少的内容，其内容标签保留，并输出提示。



## 效果展示

模板：

![template-demo_01](imgs/template-demo_01.png)

生成文档：

![word_01](imgs/word_01.png)



## 其他说明

本项目只依赖了`python-docx`第三方库。欢迎点赞收藏。
