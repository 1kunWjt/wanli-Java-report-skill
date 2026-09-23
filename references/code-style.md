# 源代码风格（参照课堂示例）

实验报告里的 Java 代码应像学生课上刚写出来、能跑通的版本：短、直接、不过度设计。

## 结构

- 每题一个完整类，含 `public static void main(String[] args)`。
- 类名简单：`t1`、`t2`、`Test1`、`Hello`、`Basic` 等。
- 允许 `class Hello`（无 public）这种实验常见写法。

## 示例形态（与老师示例一致）

```java
public class t2 {
    public static void main(String[] args) {
        int[][] k;
        k = new int[3][4];
        for (int i = 0; i < 3; i++)
            for (int j = 0; j < 4; j++)
                k[i][j] = i * j;
        System.out.println(k[2][3]);
    }
}
```

要点：

- 数组题：先声明/分配，再用 `for` 初始化，最后打印指定下标。
- 命令行输入：`args[k]` + `Integer.parseInt(args[k])`。
- 字符串题：`args.length`、`args[k].length()`。
- 判断：`if/else` 或三元运算符；奇偶用 `%`。
- 输出：`System.out.println`，字符串拼接用 `+`。

## 允许的知识点

基本数据类型、运算符、数组（一维/二维）、`args`、`Integer.parseInt`、`String.length()`、`if/else`、三元、`for`/`while`。

## 禁止

- 文件读写、Scanner 交互（除非题目明确要求）、集合、异常处理框架、类继承/接口（非 OOP 实验主题时不要硬加）。
- 过长的工具类、设计模式、英文长命名重构。
- 伪代码、残缺代码（缺类名或 main）。

## 运行说明

报告正文**不要**写「运行示例：java …」。命令行传参只体现在源代码里的 `args` / `Integer.parseInt` 即可。
