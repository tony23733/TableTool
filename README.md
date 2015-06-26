# TableTool
游戏数值表自动读取生成工具

## 如何使用
>Step 1：运行**TableTool**，将数据表格拖动到中间区域。
>Step 2：勾选**U3D(C#)**，点击**生成全部表格**，生成对应的**.bytes**文件和**.cs**文件。
>Step 3：将生成的文件拷贝到Unity工程里，用作数据。
>Step 4：将**LoadTableManager.cs**拷贝到工程里，用作读取数据。参见**Test**工程示例。

## 配置表格
>第一行是字段名，区分大小写。
>第二行表示数据类型。
>第三行是说明描述，程序忽略此行。

数据类型包含：  
整型：INT  
浮点类型：FLOAT  
字符串类型：STRING  
枚举型：ENUM  
Enumerate1=1  
Enumerate2=2  
长整型：LONG  
双精度浮点型：DOUBLE  

标记表示法：  
>关键字：星号*前缀  
>生成枚举定义：#前缀  
>忽略字段：-前缀  
**注：标记可以组合使用。**

## 生成的数据文件格式
[版本号4Byte]+[数据条数4Byte]+{[数据类型1Byte]（注：INT=1,FLOAT=2,STRING=3,ENUM=4,LONG=5,DOUBLE=6）+[域xByte]....}

## 文件结构
生成工具：  
--TableTool.sln  
--TableTool  

Unity3D读取数据脚本：  
--LoadTableManager  
  --Unity  
    --LoadTableManager.cs  

示例表格：  
--Test  
  --Table  
    --AreaMainKey.xls  
    --AreaDoubleKeys.xls  

Unity3D示例工程：  
--Test  
  --UnityProject  

## 示例
AreaMainKey.xls和AreaDoubleKeys.xls数据内容完全一样。AreaDoubleKeys是演示多个检索关键字的配置方法。对比两个表，AreaMainKey表中id为6的单元格是空的，此处将填充默认数据。

## Support
QQ群：255316030