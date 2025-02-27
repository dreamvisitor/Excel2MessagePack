# Excel2MessagePack  
A tool for converting Excel content to MessagePack files  
一个将excel内容转为messagepack文件的工具  
  
## 项目依赖  
EPPlus  
MessagePack-CSharp  
  
## 配置文件  
首次运行时，会在当前工作目录创建setting.ini文件。格式如下：  
[Settings]  
TargetFolder=Excel所在的文件夹目录  
OutputFolder=转换后的Messagepack二进制与json目录  
SourceCodeFolder=根据表头生成的C#代码目录  

## Excel文件  
Sheet名称xxx 会作为类名 class xxx 与 class xxxMgr名称  
第1行 A1单元格配置数据容器类型：map 或者 list  
第2行 属性名称  
第3行 数据类型 支持(int,double,bool,long,short,float,string) 默认为string  
第4行 字段说明
第5行... 数据  
![image](https://github.com/user-attachments/assets/79fedf7e-bcbb-4d94-bf36-9139a0084e41) 
![image](https://github.com/user-attachments/assets/9cc072e4-5eaa-4be6-8019-213b900c4820)

## Tips
本工程为.net 9.0  
请更新后自行编译exe文件  

## LICENSE
MIT
