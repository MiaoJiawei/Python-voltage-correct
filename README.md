# Python-voltage-correct

## 上手指南  
变扫速CV的电压校正程序（顺带拟合）  
使用方式非常简单粗暴，只需要在运行目录底下放一个raw.xlsx文件  
格式为第一行的偶数位填扫速  
从第四行开始填写数据，奇数列填电压，偶数列填电流  
然后运行，就可以在源目录产生res.xlsx文件  

### 使用要求
环境：*Python 3*  
需要的模块：numpy, scipy, openpyxl  
> *matplotlib 可选，调试的时候暂时希望能即时输出，等完事了会去掉*

## 后记  
本程序理论指导为《电化学原理 方法与应用》，挺实用的，力荐  
本程序思路来自曹余良老师的一个讲座  
感谢新哥、韩熙、李佳鹏在饭点喊我吃饭  
感谢鹏姐、曾经的上铺在测试数据方面提供的便利  
