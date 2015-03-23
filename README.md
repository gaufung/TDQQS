# TDQQS
全国农村土地确权内业处理程序
## 前言
本程序采用的是ArcGIS Object(AO)二次开发，采用C#语言开发的desktop应用程序，主要作用是针对农业部关于农村[土地确权](http://www.mlr.gov.cn/zwgk/zytz/201105/t20110516_865762.htm)的颁证工作的展开
##系统要求
+ Win操作系统
+ .Net framework3.5或者以上
+ ArcGIS Desktop10.2 版本或者以上
+ Office 环境

## 数据
+ 地图数据库
数据采集工作是采用野外GPS实测田块边界点，内业使用ArcMap 采用数字化地块的方式，通过两点连线，线构面成田块，并输入相应的承包方名称（CBFMC）和合同面积（HTMJ），并添加其余的字段。数据传输方式采用个人地理数据库，即Access数据库。
+ 基础数据库
在工作展开之前，先录入了每个行政村的基础信息，承包方信息，家庭成员信息，并提供了基础数据库模板（MDB文件）。两者数据库在同一个文件夹目录下，方便程序读取。

## 软件设计
* WPF桌面应用程序开发
* WPF程序中MVVM设计模式
* 工厂模式
