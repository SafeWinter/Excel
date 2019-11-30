# Excel
Share and store some Excel VBA code



---

> VBA_tutorial folder
>
> Learing notes on ExcelHome's VBA course -- *Learning Excel VBA from scratch* (May, 2017)

> VBA_advanced folder
>
> Learing notes on ExcelHome's VBA course -- *Learning Excel VBA in Action* (Aug, 2017)



## VBA 培训知识点摘要

### 《VBA 零基础》笔记

> **L01：开篇**
>
> 讲述了两个宏录制案例：工资条制作与文本分列操作，宏绑定、宏安全性设置、VBA 版 `HelloWorld` 程序、代码注释、结构化语言三大结构、标识符命名规范

> **L02：变量（第一部分）**
>
> *课程*——
>
> 变量的定义、声明方法（总统套房）、显式声明与隐式声明、变量重复声明问题、常见的数值型数据类型（Integer、Long、Double）、VBA 四舍五入规则、常见数据类型的简写形式（%&$!#@）、常见运算符、表达式与运算优先级、混合数据类型运算规则（自动强制转换）
>
> *作业*——
>
> 计算平均年龄时用到了格式化函数 `Format(exp, pattern)`
>
> 计算出生天数用到了日期转换函数 `CDate(exp)`



### 《VBA 实战进阶》笔记

> **L01：再谈变量**
>
> *课程*——
>
> 　　介绍了局部自动变量、局部静态变量、模块级变量、全局变量、变量使用原则、变量阴影效应的相关概念，并从编译原理的角度讲解了计算机内存的四个区域：代码区、数据区、栈区、堆区
>
> *作业*——
>
> 1. 教室一周上课情况高亮展示
> 2. 全年级学生成绩模糊查询
>
> *代码*——
>
> 1. 工作表改变事件用到了 `Intersect` 方法
> 2. 查询列的动态选择用到了 `IIF` 函数
> 3. 关键词类型的判定用到了 `IsNumeric(keyword)` 函数
> 4. 初始化清空单元格：`resultSheet.UsedRange.Offset(1, 1).ClearContents`