2020-07-12
做了添加：
1.后面因为每个月都要手动复制上个月的信息，懒得搞，又写了个方法，可以自动复制上个月的信息，修改为当前/下一个月
2.因为要算水电损耗多少，还得统计该栋房子所以租客的水电情况，核算与总表之间的差额，所以又搞了个表格，清晰看出每个租客水电情况，还有合计

至于图，懒得贴了。。


=================================================================================================================================

在微信上给租客发送房租

这是利用python写成的发送房租软件  
原理：用tkinter写出界面，打开记录房租的excel文件，获取数据后再生成图片，在界面上显示所有房租信息，接着可以登录微信，给所有租客发送其该月的房租  

需要的条件：  
1.记录房租的excel文件  
2.生成图片的excel文件  
3.excel软件  
4.微信上备注好租客昵称  
5.python各种包，比如win32com、openpyxl、wxpy  

注意事项：  
1.房租文件名称样式：文件名为房屋名，sheet为房间号，例：其对应的租客昵称为："一楼-1房"，（见附带文件）  
2.房租文件里的各项信息的样式(见附带文件)，可以自己修改，但要一一对应生成图片的excel文件--“截图.xlsx”  
3.登录微信，实际上是用wxpy登录网页版，建议自己先试试其微信能不能登陆，如果不能，换号吧  

操作步骤：  
运行程序后，  
1.先点击“获取房租信息”按钮，所有租客该月的房租信息会显示在treeview上（如果该租客此月房租还没记录好，那么其对应的按钮不可选），并依据其房租信息生成表格图片。  
2.点击“登录微信”，弹出二维码，扫描登录后，按钮变为“发送房租”，此时可以选择要发送的租客，并且设定好微信昵称类型(房间名就是“楼房-房号”），再点击发送即可  
3.发送完毕后，每个租客的发送结果会显示在tree上  

微信上发送的内容：  
--登录微信、发送了信息后，会有提示发送到自己微信的文件助手  
--发送给租客内容为2部分：1.租金；2.详细租金表格(水费电费等等）  

打包为exe：  
因为需要操作excel，好像只能在windows上运行  
打包命令：pyinstaller -F begin.py -w （无命令窗口，并且压缩为单个exe）  
(写了个日志输入，error信息会打印在“日志.log”文件上）  

详细解说见：https://blog.csdn.net/qq_38282706/column/info/41792

![image](https://github.com/zhishiluguoliu6/Send-rent-to-tenants-on-Wechat-/blob/master/%E7%A4%BA%E4%BE%8B%E5%9B%BE%E7%89%87/%E5%8F%91%E9%80%81%E4%BF%A1%E6%81%AF.jpg)
