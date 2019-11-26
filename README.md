# 报告自动校核
基于WPF的报告自动校核软件，可以校核单位模板报告中的一系列错误，目前主要针对桥梁荷载试验及桥梁常规定期检测的报告

## 软件使用方法
![软件主界面](https://github.com/zjkl19/WPFAIReportCheck/blob/master/软件主界面.png)

1、运行“WPFAIReportCheck.exe”；

2、根据UI提示：选择需要校核的报告=>开始校核=>查看校核汇总结果=>查看标出问题的报告

## 软件目前可以校核的问题：
1、查找单位描述错误。如km/h写成Km/h；

2、查找规范描述错误。如规范中某个字母打错，某个地方多写/少写了空格（正确的规范名称库在程序根目录下"规范.xlsx"文件中）；

3、查看报告是否对构件编号进行了说明；

4、名词描述错误。如“CNAS检验机构”写成“CNAS检查机构”（名词描述错误数据库在程序根目录下"描述错误.xlsx"文件中）；

5、查找报告中是否存在其它桥梁。如全文描述的是福州xx桥，但是文中某处却出现漳州某桥；

6、查找页码对应错误。如目录中附录在第17页，实际却在第18页；

7、查找表格中序号是否连续；

8、验算桥梁静力荷载试验中，应变和挠度表格中各个参数的计算结果；

9、查看桥梁静力荷载试验中，应变、挠度的统计结果和报告汇总表格中是否一致；

10、查看是否存在图片和对应描述跨页的情况；

11、查看是否存在表格和对应描述跨页的情况；

12、查看表格标题中序号是否存在不连续的状况（仅支持类似“表 x-y”格式，不支持类似“表1”，“表3.4”等格式）

13、查看图片标题中序号是否存在不连续的状况（仅支持类似“图 x-y”格式，不支持类似“图1”，“图3.4”等格式）

14、搜索全文是否含有半角逗号，若搜到，提示是否需要修改
