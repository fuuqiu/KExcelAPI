# KExcelAPI [![Build Status](https://travis-ci.org/webarata3/KExcelAPI.svg?branch=master)](https://travis-ci.org/webarata3/KExcelAPI) [![Coverage Status](https://coveralls.io/repos/webarata3/KExcelAPI/badge.svg?branch=master&service=github)](https://coveralls.io/github/webarata3/KExcelAPI?branch=master)
它是Kotlin的Apache POI包装器。, 我想让每个人都喜欢Excel尽可能简单。, 该库是在[GExcelAPI]（https://github.com/nobeans/gexcelapi）的影响下创建的。
## How to use
从工作表中，您可以通过指定单元名称或单元格索引来检索Cell对象。
 由于我们需要一个类型来从单元格对象中检索数据，
 我们将使用toInt，toDouble，toStr等方法获取值。
 对于一组数据，类型很明显，因此您可以设置值, 带“=”的Excel。
 作为一种基本使用方法，我们导入`link.webarata3.kexcelapi。 , *`
 包因为我们正在使用扩展功能。
```kotlin
import link.webarata3.kexcelapi.*
```
仅此就完成了准备工作。, 您现在可以按如下方式访问Excel。
```kotlin
// Easy file open, close
KExcel.open("file/book1.xlsx").use { workbook ->
    val sheet = workbook[0]

    // Cell loading
    // Access by Cell name
    println("""B7=${sheet["B7"].toStr()}""")
    // Access by Cell index [x, y]
    println("B7=${sheet[1, 6].toDouble()}")
    println("B7=${sheet[1, 6].toInt()}")

    // Write Cell
    sheet["A1"] = "ABCDE" // Japanese
    sheet[3, 7] = 123

    // Easy to write files
    KExcel.write(workbook, "file/book2.xlsx")
}
```

## Maven

The Maven repository (Gradle setting) is as follows.

```groovy
repositories {
    maven { url 'https://raw.githubusercontent.com/webarata3/maven/master/repository' }
}

dependencies {
    compile 'link.webarata3.kexcelapi:kexcelapi:0.5.0'
}
```

#Fuuqiu Version
- 非常感谢作者提供相对家简洁的操作Excel的一套API

- 修改说明
  - 1.因本人项目中文件传输使用流的形式,需要在操作Excel时业务代码中已经获取到流,如果使用原生提供的API需要存储文件,后以文件形式读取到项目.造成一定的繁琐流程,故直接修改源码打成私包上传至Nexus,项目引入
  - 2.顺便修改了一下源码中日语的注释和提示信息
  




## License
MIT
