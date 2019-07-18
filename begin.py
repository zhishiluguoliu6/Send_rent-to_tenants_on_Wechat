

from PIL import ImageGrab,Image
from win32com.client import Dispatch, DispatchEx
import pythoncom,os,itertools,time
from collections import OrderedDict
import openpyxl,gc

import logging,time,traceback,datetime
from wxpy import *
from tkinter import *
from tkinter.ttk import *
#设定log的输出设置
logging.basicConfig(level=logging.WARNING,
                    format='asctime:        %(asctime)s \n'  # 时间
                           'bug_line:       line:%(lineno)d \n'  # 文件名_行号
                           'level:          %(levelname)s \n'  # log级别
                           'message:        %(message)s \n',  # log信息
                    datefmt='%a, %d %b %Y %H:%M:%S',
                    filename='日志.log',  # sys.path[1]获取当前的工作路径
                    filemode='a')  # 如果模式为'a'，则为续写（不会抹掉之前的log）

class Open_Excel():
    '''打开当前文件夹所有xlsx文件'''
    def __init__(self,month):
        self.all_info=OrderedDict()# 每个住户对应其数据组成的dict,格式{住户A:{X月份:{电费:xx,合计:xx},y月份:....},住户B....}
        self.month=month  #tkinter上选择的月份

    #
    def file_list(self):
        '''获取当前文件夹内所有出租房的xlsx文件 组成list,
            (不是xlsx文件、截图文件、打开状态的excel文件)'''
        file_list=[]
        for file in os.listdir('.'):
            if os.path.splitext(file)[1]=='.xlsx' and file!='截图.xlsx' and '~$'not in file:
                file_list.append(file)
        return file_list


    def get_excel_info(self, file):
        '''循环每个excel文件里的sheet，获取每个住户对应的所有租金详情'''
        wb = openpyxl.load_workbook(file,data_only=True) #打开excel，data_only取excel显示的值而不是公式

        # 循环每个sheet
        for sheet in wb.worksheets:
            sheet_value =list(sheet.values)  # 当前sheet的所有value
            keys = sheet_value[0] # ['电表', '水表', '用电量', '用水量', '电费', '水费', '房租', '垃圾费', '其他', '合计']
            sheet_dict = {}  # 存放每个sheet里的数据，键为每个月

            # 循环每行(每个月)
            for month_value in sheet_value[1:]:
                if month_value[11] == None: continue  # 合计为空的跳过
                month_dict = {}  # 每个月里的每一项组成的dict
                # 循环每个项目，放进每个月的数据dict
                for i in range(len(month_value)):
                    key=keys[i]
                    value = month_value[i]  # 每个项目里的数据
                    month_dict[key] = value

                the_month = '%s-%s' % (month_value[0], month_value[1])
                sheet_dict[the_month] = month_dict  # 每个月对应的详细 租金

            the_zuhu = os.path.splitext(file)[0] + '-' + sheet.title  # 住户名由 房子名+房号
            self.all_info[the_zuhu] = sheet_dict  # 所有 住户 组成的dict

        wb.close()      # 关闭Excel文件，不保存
        del wb#删除工作簿
        gc.collect()#内存释放

    #
    def save_img(self,file,month_info,send_info):
        '''根据每个住户对应的租金，在截图xlsx文件修改数值，然后截图保存
        :param file: 截图.xlsx
        :param month_info: 所有住户此月的租金信息dict
        :param send_info: 全部住户的租金信息(用于微信发送)
        :return:
        '''
        file_name = os.path.abspath(file)  # 把相对路径转成绝对路径
        pythoncom.CoInitialize()  # 开启多线程
        excel = DispatchEx('excel.application')# 创建Excel对象
        excel.visible = False         # 不显示Excel
        excel.DisplayAlerts = 0     # 关闭系统警告(保存时不会弹出窗口)

        workbook = excel.workbooks.Open(file_name)# 打开截图.xlsx
        wSheet = workbook.worksheets['截图']

        # 循环每个住户的数据,根据 所选月份，得到具体数据，然后再截图xlxs上改变数字，截图保存，添加到send_info
        for the_zuhu,the_month_data in month_info.items():
            #该住户此月的租金信息不为空
            if type(the_month_data)!=int:
                img_name = self.month + '：' + the_zuhu  #该住户此月的截图名
                self.change_sheet(wSheet, the_month_data)   #根据不同住户 改变截图xlsx 里的 每个项目的金额
                self.snapshot(excel, wSheet, img_name)#截图，保存
                send_info[the_zuhu] = [the_month_data['租户'],the_month_data['合计'], img_name + '.png'] #格式：{住户A：[租户名，合计租金，图片名称],住户B....}
            else:
                send_info[the_zuhu]=0

        workbook.Close(False)  # 关闭Excel文件，不保存
        excel.Quit()  # 退出Excel
        pythoncom.CoUninitialize()  # 关闭多线程


    def change_sheet(self,sheet,data):
        '''根据不同住户 改变截图xlsx  每个项目的金额'''
        the_key = data.keys()  # [月份、租户，'电表', '水表', '用电量', '用水量', '电费', '水费', '房租', '垃圾费', '其他', '合计']
        all_range = itertools.chain(sheet.usedrange)  # 合并所有单元格元素

        # 循环所有单元格，如果是电费、水费等项目，就修改其对应的单元格数值
        for one in all_range:
            if one.value in the_key:
                one.offset(1, 2).value = data[one.value]

    def snapshot(self,excel,sheet,img_name):
        '''
        :param excel: win32 的excel对象
        :param sheet: 截图sheet
        :param img_name: 月份+住户
        :return:
        '''
        # 选定截图区域，保存img文件
        sheet.UsedRange.CopyPicture()  # 复制有内容的单元格区域
        sheet.Paste()  # 粘贴
        excel.Selection.ShapeRange.Name = img_name  # 将刚刚选择的Shape重命名，避免与已有图片混淆
        sheet.Shapes(img_name).Copy()  # 选择图片
        img = ImageGrab.grabclipboard()  # 获取剪贴板的图片数据
        img.save(img_name + ".png")


    def get_all_info(self,):
        '''获取所有住户的全部租金信息，放到self.all_info
           根据选取的月份，筛选后添加到month_info返回'''

        #循环当前文件夹的所有excel文件，获取每个住户对应的所有具体租金，存放到self.all_info
        for file in self.file_list():
            self.get_excel_info(file)

        month_info=OrderedDict()
        #循环所有住户、所有月份的 dict
        for the_zuhu,sheet_dict in self.all_info.items():
            the_month_data=sheet_dict.get(self.month)    #提取tk上所选月份 的具体租金dict
            #该住户有此月份的 租金信息，那么就添加到 month_info，否则为0
            if the_month_data:
                the_month_data['月份']=self.month
                month_info[the_zuhu]=the_month_data
            else:
                month_info[the_zuhu] = 0
        return month_info

    def get_send_info(self):
        time1=time.time()
        month_info=self.get_all_info()  #获取所有 租户 所选月份对应的所有具体租金
        print(time.time()-time1) #1.443082571029663
        send_info = OrderedDict()  # 在微信发送的dict，格式：{住户A：[租户名，合计租金，图片名称],住户B....}
        time1 = time.time()
        self.save_img('截图.xlsx', month_info,send_info) #将房租信息处理成图片，并把对应信息放入send_info
        print(time.time() - time1) #3.7622151374816895
        print(send_info)
        return send_info


'''Open_Excel是根据每个excel里的每个租客租金详情，生成房租信息send_info 以及对应的表格图片
   过程：Open_Excel(月份) 输入月份实例化
         get_send_info() 运行
         get_all_info    获取当月所有租户具体租金
                        ---get_excel_info(file) 打开每个excel获取所有房租信息 （运用了openpyxl）
                        ---month_info   存放 目的房租信息
         save_img()     将房租信息处理成图片，并把对应信息放入send_info   （运用了win32com）
                        -循环所有租户：
                                     ---change_sheet()   #根据不同住户 改变截图xlsx 里的 每个项目的金额
                                     ---snapshot()#截图，保存
                                     ---send_info 保存格式：{住户A：[租户名，合计租金，图片名称],住户B....}
         '''





'''tk窗口
   1.年月时间选择窗口
        要点：根据当前时间自动选择好
             选择月份，会在text窗口提示
             
   2.text显示操作窗口
        要点：不同操作，字体不同————初始、常规、警告、报错
                     1.选择时间                              常规
                     2.选择微信昵称                           常规
                     3.获取xx月所有房租                       常规 
                     4.报错：获取xx月房租失败，再按一次        报错 
                     5.登录微信《》成功                       常规 
                     6.报错：微信登录失败                     报错  
                     7.点击发送：① 所有住户都没有房租信息         警告
                                ②此次发送的是房租/特定内容       常规             
                                ③报错：此时已掉线，请重新登录    报错                                
                                ④没有选中租户，请选择            警告
                                ⑤成功，发送了N个租户             常规
              插入后，text窗口不可编辑，而且一直显示的是最新内容
              
     发送窗口
        构成：1.租客在微信上的昵称 单选按钮(注：房间名为 “房屋-房号” 组成)  
              2.发送特定内容(输入框/登录-发送按钮)
   
   3.主要操作的按钮
        构成：获取房租信息：
                  步骤：清空tree
                        调用Open_Excel，获取住户租金详情，生成截图
                        将租金详情在tree上显示出来，将按钮、截图等关系放在self.orm （具体见insert_tv()）
              登录微信/发送房租：
                  ————登录：使用wxpy登录微信，成功后改变按钮状态
                  ————发送：
                          1.判断是否获取了住户详情
                          2.判断所选月份下，获取的住户能否被选中
                          3.判断此时微信是否在线：
                          4.循环self.orm(存放每个住户信息的dict)
                                发送条件：---此住户被选中(多选按钮打勾)
                                         ---微信有此住户
                          5.text插入发送情况
    
   4.tree显示窗口：
        构成：——tree头，因为tree原始的头会随滚动条动，所以用按钮重设了一个头
             —— 主体： 滚动条
                       画布canvas：tree
                                   多选按钮frame
        要点：鼠标滚轮滚动时，改变的页面是canvas整个画面(包括多选按钮) 而不是单独treeview
              设定tree的样式：---每行高度
                             ---颜色：常规、被选中时、不可选时
    
   5.insert_tv 插入tree时 设定每个item跟多选按钮 的联系
        插入：
               循环send_info_dict：---如果value是数字(0)，那么此租户没结算房租
                                  ---结算了房租：tree插入该租户
                                                将该有效的租户信息放入self.orm
                                  ---创建多选按钮，与tree该行(item)绑定
        要点：
               根据tree重设定窗口tv_frame的高度
               多选按钮与item绑定：勾选按钮/点击item，2者都会发生改变，
               每次选择后，全选按钮也会发生改变
               没有结算房租的，其状态为不可选
               
'''

class My_Tk():
    def __init__(self):
        self.tk=Tk()
        self.tk.geometry('665x600')
        self.tk.resizable(width=False, height=True)  # 宽不可变, 高可变,默认为True

        self.create_yearframe()     #创建年份 构件
        self.create_monthframe()    #创建月份构件
        self.create_stateframe()    #创建text构件，包含发送特殊内容按钮
        self.create_buttonframe()   #创建 获取房租信息、登录微信 按钮
        self.orm = {}               #获取房租信息后，存放tree每一行、多选按钮、与租客的对应dict {item:按钮、住户房号、租客名、image}

        self.create_heading()       #tree重设的头
        self.create_tv()            #创建tree与多选按钮 构件
        mainloop()


# =============================最开始选择年月部分========================###
    def create_yearframe(self):
        '''创建年份的构件，放入单选按钮'''
        import tkinter
        yearframe=tkinter.LabelFrame(self.tk,height=50, width=400, text='年份',)
        yearframe.pack(fill=X)

        self.year = StringVar()
        year=datetime.datetime.now().strftime('%Y')
        self.year.set(year)

        Style().configure('TRadiobutton', font='宋体', )
        for i in range(2019,2023):
            month=Radiobutton(yearframe,variable=self.year, text='%s年'%i, value='%s'%i)
            month.grid(column=i, row=1, sticky=W, padx=10)

    def create_monthframe(self):
        '''创建月份的构件，放入单选按钮，只有选择月份时，才有调用回调函数，显示所选时间'''
        import tkinter
        monthframe=tkinter.LabelFrame(self.tk,height=50, width=400, text='月份')
        monthframe.pack(fill=X)

        self.month = StringVar()
        month=datetime.datetime.now().strftime('%m')
        self.month.set('%s月'%int(month))
        def select_month():
            self.the_month = '%s-%s' % (self.year.get(), self.month.get())  # 获取tk上选择的时间
            self.text_insert("选择时间为： %s" % self.the_month)

        for i in range(1,7):
            month=Radiobutton(monthframe,variable=self.month, text='%s月'%i, value='%s月'%i,command=select_month)
            month.grid(column=i, row=1, sticky=W, padx=20)

        for i in range(7, 13):
            month=Radiobutton(monthframe,variable=self.month, text='%s月' % i, value='%s月' % i,command=select_month )
            month.grid(column=i-6, row=2, sticky=W,padx=20)
# =============================最开始选择年月部分========================###




# ==================================text等的中间部分========================###
    def create_stateframe(self):
        '''
         此构件分2个部分：
            一：text部分，每次按动按钮，其操作会显示在text上；
            二：微信发送部分，分为：1.选择发送昵称(微信上备注)---房间名  ---租户名
                                 2.发送特定内容
                                '''
        #整个构件
        stateframe=Frame(self.tk)
        stateframe.pack(fill=BOTH)

        #======text、滚动条 构件创建====
        textframe=Frame(stateframe)
        textframe.pack(fill=BOTH,side=LEFT)

        self.state_text = Text(textframe, width=60,height=7, )
        bar = Scrollbar(textframe)
        # 两个绑定
        bar.config(command=self.state_text.yview)
        self.state_text.config(yscrollcommand=bar.set, )
        # 固定位置
        bar.pack(side=LEFT, fill=Y)
        self.state_text.pack(side=LEFT, fill=X, expand=1)

        #设定各种text字体格式
        #分别有初始、常规、警告、报错
        self.state_text.tag_config('start', font=('system', 12, 'bold'),background='DarkGray', foreground='Chartreuse')
        self.state_text.tag_config('default',font=('system',12))
        self.state_text.tag_config('warning',foreground='blue', font=('system',12,'bold'))
        self.state_text.tag_config('error',foreground='red', font=('Fixdsys',12),underline=True,background='Wheat')
        self.text_insert("欢迎使用微信发送房租软件，请先选择月份",'start')


        #======发送构件===================================
        wxframe = Frame(stateframe)
        wxframe.pack(fill=BOTH,side=LEFT,padx=10,)

        #发送昵称 构件
        name_frame=LabelFrame(wxframe,text='微信昵称')
        name_frame.pack(fill=BOTH,padx=10, )

        self.send_name = StringVar()
        self.send_name.set('房间名')
        Style().configure('W.TRadiobutton', font='宋体 11', )
        Radiobutton(name_frame, variable=self.send_name, text='房间名', value='房间名', command=self.set_sendname,style='W.TRadiobutton').pack(side=LEFT)
        Radiobutton(name_frame, variable=self.send_name, text='租客名', value='租客名',command=self.set_sendname,style='W.TRadiobutton' ).pack(side=LEFT)


        # 微信发送特定内容 构件
        sendframe = LabelFrame(wxframe, text='发送信息')
        sendframe.pack(fill=BOTH, padx=10, )

        self.word_var=StringVar()
        send_word=Entry(sendframe, textvariable=self.word_var)
        send_word.pack()
        self.word_button=Button(sendframe,text='登录微信',command=lambda :self.log_wx())
        self.word_button.pack()


    def text_insert(self,text,tags='default'):
        '''在text中 根据不同的样式 插入各种信息'''
        self.state_text['state']=NORMAL
        self.state_text.insert("end", '\n·')
        self.state_text.insert("end",text,tags)
        self.state_text.yview_moveto(1)
        self.state_text['state'] = DISABLED #让text栏不可修改

    def set_sendname(self):
        '''选择微信昵称的回调函数'''
        send_name=self.send_name.get()
        self.text_insert("选择微信昵称为： %s" % send_name)
#==================================text等的中间部分========================###



# ==================================按钮部分，各种回调函数========================###
    def create_buttonframe(self):
        the_frame=Frame(self.tk)
        the_frame.pack(fill=X)
        Label(the_frame,text='全选').pack(side=LEFT)
        self.get_info_button=Button(the_frame,text='获取房租信息',width=25,command=self.open_excel)
        self.log_wx_button = Button(the_frame,text='登录微信',width=25,command=lambda :self.log_wx())

        self.get_info_button.pack(side=LEFT)
        self.log_wx_button.pack(side=LEFT,padx=20)


    def open_excel(self):
        '''回调函数，====获取房租信息
           清除tree里的内容
           根据所选月份，调用Open_Excel，获取所有房租信息，生成截图；
           调用insert_tv将该月份的租金详情插入tree'''
        self.clear_tv()                     #删除tree
        self.the_month = '%s-%s' % (self.year.get(), self.month.get())  # 获取tk上选择的时间
        self.get_info_button['state'] = 'disabled'
        self.text_insert("获取<= %s =>的房租信息"%self.the_month)
        self.tk.update()
        try:
            open_excel=Open_Excel(self.the_month) #根据所选月份，实例化打开所有excel的类
            send_info_dict=open_excel.get_send_info()#获取所有住户 该月份的租金信息，截图
            self.insert_tv(send_info_dict)          #将住户的租金信息 插入到tree
        except Exception as e:
            self.text_insert("获取<= %s =>的房租信息失败，请再按一次！！！"%self.the_month, 'error')
            logging.error("ERROR：%s\n"
                      "%s" % (e, traceback.format_exc()))
        finally:
            self.get_info_button['state'] = 'normal'


    def log_wx(self,cache_path=True):
        '''回调函数，
           登录微信，登录成功，改变2个按钮状态，重设回调函数'''
        try:
            self.bot = Bot(cache_path=cache_path)
            time.sleep(0.7)
            self.bot.file_helper.send('登录发送房租软件')
        except Exception as e:
            self.text_insert("登录微信失败，请重新登录！！" , 'error')
            logging.error("ERROR：%s\n"
                          "%s" % (e, traceback.format_exc()))

            self.log_wx_button['command'] = lambda: self.log_wx(False)  # 登录微信不用缓存
            self.word_button['command'] = lambda: self.log_wx(False)  # 登录微信不用缓存
        else:
            self.text_insert("登录微信 %s 成功" % self.bot.self)
            #成功登录微信后，改变按钮状态
            self.log_wx_button['text']='发送房租'
            self.log_wx_button['command']=lambda :self.wx_send(self.send_fangzu) #此时按钮变为发送房租

            self.word_button['text'] = '发送信息'
            self.word_button['command'] = lambda: self.wx_send(self.send_words)  # 此时按钮变为发送特定信息



    def wx_send(self,send_func):
        '''回调函数，登录微信后，根据不同按钮，发送房租/特定信息
           步骤：1.判断是否获取了住户详情：没有，那就按下【获取房租信息】，中断
                2.判断所选月份下，获取的住户能否被选中：都不能选中，text插入警告，中断
                3.判断此时微信是否在线：---在线，给自己发送信息；text插入信息
                                     ---掉线，text插入报错，并log
                                         改变按钮状态，然后中断
                4.循环self.orm(存放每个住户信息的dict)
                  发送条件：---此住户被选中(多选按钮打勾)
                           ---微信有此住户
                  根据send_func  发送房租信息/特定信息
                5.text插入发送情况

        '''
        #还没有获取住户详情，那么就按下【获取房租信息】
        if len(self.tv.get_children())==0:
            self.open_excel()
            return

        #此时所有住户都没有房租详情(不能选择发送)
        if len(self.orm)==0:
            self.text_insert("<= %s =>所有住户都没有房租信息，请重选"%self.the_month, 'warning')
            return

        #测试微信是否还在线，若不在，更改按键信息，报错，中断发送
        log_word='发送<= %s =>房租信息'%self.the_month
        if send_func==self.send_words:log_word='发送特定信息：%s'%self.word_var.get() #不同按钮，内容不同
        try:
            self.bot.file_helper.send(log_word)
            self.text_insert(">>>>>>%s" % log_word)
        except Exception as e:
            self.text_insert("微信登录失败，请重新登录！！！" , 'error')
            logging.error("ERROR：%s\n"
                          "%s" % (e, traceback.format_exc()))

            self.log_wx_button['text'] = '登录微信'
            self.word_button['text'] = '登录微信'
            self.log_wx_button['command'] = lambda :self.log_wx(False) #登录微信不用缓存
            self.word_button['command'] = lambda: self.log_wx(False)  # 登录微信不用缓存
            return

        # 循环orm里每个住户的信息，若此住户被选中，而且在微信好友中，那么发送信息send_info:[总房租、截图]
        send_nums=0             #每次发送用户的数量
        wx_friends=self.bot.friends()                       #微信的好友列表
        for item,[button,zuhu,name,image] in self.orm.items():
            if self.send_name.get()=='租客名':zuhu=name
            friends = wx_friends.search(zuhu)              #根据该住户获取微信好友(列表)
            button_value = button.getvar(button['variable']) #获取对应按钮状态
            #当前住户是被选中(打勾)
            if button_value == '1':
                button.invoke()             #按钮变为不选中
                #微信添加了该租户
                if len(friends) > 0:
                    send_nums += 1
                    send_func(friends, item, name, image)
                else:
                    zhuangtai = '没有添加此住户的微信'
                    self.tv.set(item, column='发送详情', value=zhuangtai)

        #每次发送后，text插入发送情况
        if send_nums==0:
            self.text_insert("请选择要发送的住户!!!" , 'warning')
        else:
            self.text_insert("→→→→→→成功发送了  %s 个住户\n"%send_nums)


    def send_fangzu(self,friends,item,name,image):
        '''给各个住户 发送【房租信息】'''

        # 微信添加了该住户(备注好了)
        the_zuhu = friends[0]  # 住户的微信号

        fangzu = self.tv.set(item, column='房租')  # 获取房租金额
        the_zuhu.send('%s，您%s的房租是：%s元' % (name,self.month.get(),fangzu))
        time.sleep(0.5)
        the_zuhu.send_image(image)  # 发送截图
        time.sleep(0.5)
        zhuangtai = '房租发送成功'
        if len(friends) > 1: zhuangtai = '房租发送成功(微信号不止1个)'
        self.tv.set(item, column='发送详情', value=zhuangtai)  # tree上更新发送详情



    def send_words(self,friends,item,name,image):
        '''给各个住户 发送【指定信息】'''
        the_zuhu = friends[0]  # 住户的微信号

        words=self.word_var.get()
        the_zuhu.send(words)
        zhuangtai = '特定内容发送成功'
        if len(friends) > 1: zhuangtai = '特定内容发送成功(微信号不止1个)'
        self.tv.set(item, column='发送详情', value=zhuangtai)  # tree上更新发送详情
# ==================================按钮部分，各种回调函数========================###


# ==================================tree部分，构建========================###
    def create_heading(self,):
        '''重新做一个treeview的头，不然滚动滚动条，看不到原先的头！！！'''
        heading_frame=Frame(self.tk)
        heading_frame.pack(fill=X)

        #填充用label
        button_frame=Label(heading_frame,width=0.5)
        button_frame.pack(side=LEFT,)
        #全选按钮
        self.all_buttonvar = IntVar()
        self.all_button = Checkbutton(heading_frame, text='',variable=self.all_buttonvar, command=self.select_all)
        self.all_button.pack(side=LEFT)
        self.all_buttonvar.set(0)

        self.columns = ['日期', '楼房', '房号', '租客','房租', '发送详情']
        self.widths = [80, 70, 70, 70,70, 190]

        #用按钮作为 tree的头
        Style().configure('w.TButton', font='system',foreground='Gray')
        widths=[10, 9, 9, 9,9, 24]
        for i in range(len(self.columns)):
            Button(heading_frame,text=self.columns[i],width=widths[i]
                  ,style='w.TButton',).pack(side=LEFT)


    def create_tv(self):
        '''创建tree与多选按钮的构件
           结构：canvas_frame：---ysb(滚动条)
                              ---self.canvas：-self.tv_frame：---self.tv
                                                              ---self.button_frame：Checkbutton'''
        #放置 canvas、滚动条的frame
        canvas_frame=Frame(self.tk,width=600,height=400)
        canvas_frame.pack(fill=X)

        #只剩Canvas可以放置treeview和按钮，并且跟滚动条配合
        self.canvas=Canvas(canvas_frame,width=620,height=500,scrollregion=(0,0,620,400))
        self.canvas.pack(side=LEFT,fill=BOTH,expand=1)
        #滚动条
        ysb = Scrollbar(canvas_frame, orient=VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=ysb.set)
        ysb.pack(side=RIGHT, fill=Y)
        #!!!!=======重点：鼠标滚轮滚动时，改变的页面是canvas整个画面(包括多选按钮) 而不是单独treeview
        self.bind_mouse(canvas_frame)


        #想要滚动条起效，得在canvas创建一个windows(frame)！！
        tv_frame=Frame(self.canvas)     #注：这个是放置容器，下面带self是设定高度尺寸等
        self.tv_frame=self.canvas.create_window(0, 0, window=tv_frame, anchor='nw',width=650,height=400)#anchor该窗口在canvas左上方

        #放置button的frame
        self.button_frame=Frame(tv_frame)
        self.button_frame.pack(side=LEFT, fill=Y)
        Label(self.button_frame,width=3).pack()  #填充用


        #创建treeview==================
        self.tv = Treeview(tv_frame, height=10, columns=self.columns, show='headings')#height好像设定不了行数，实际由插入的行数决定
        self.tv.pack(expand=1, side=LEFT, fill=BOTH)
        #设定每一列的属性
        for i in range(len(self.columns)):
            self.tv.column(self.columns[i], width=self.widths[i], minwidth=self.widths[i], anchor='center', stretch=True)


        #设定treeview格式
        self.tv.tag_configure('oddrow', font='Symbol 12')                     #行的默认规格
        self.tv.tag_configure('select', background='SkyBlue',font='Symbol 12')#被选中的行背景颜色
        self.tv.tag_configure('disabled', background='Silver', font='Symbol 12') #不可选的行 背景颜色
        self.rowheight=27                                       #很蛋疼，好像tkinter里只能用整数！
        Style().configure('Treeview', rowheight=self.rowheight)      #设定每一行的高度

        # 设定选中的每一行字体颜色、背景颜色 (被选中时，没有变化)
        Style().map("Treeview",
                  foreground=[ ('focus', 'black'), ],
                  background=[ ('active', 'white')]
                  )
        self.tv.bind('<<TreeviewSelect>>', self.select_tree) #绑定tree选中时的回调函数

    def bind_mouse(self, frame):
        '''绑定鼠标滚轮与tree整个构件
           鼠标进入canvas时，滚轮滚动的是整个画面，包括多选按钮'''

        def bound_to_mousewheel(event):
            self.canvas.bind_all("<MouseWheel>",
                                 lambda event: self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))

        def unbound_to_mousewheel(event):
            self.canvas.unbind_all("<MouseWheel>")

        frame.bind('<Enter>', bound_to_mousewheel)
        frame.bind('<Leave>', unbound_to_mousewheel)

# ==================================tree部分，构建========================###


# ==================================tree与多选按钮的插入等各种操作========================###
    def clear_tv(self):
        # 清空tree、checkbutton
        items = self.tv.get_children()
        [self.tv.delete(item) for item in items]

        for child in self.button_frame.winfo_children()[1:]:  # 第一个构件是label，所以忽略
            child.destroy()
        self.tk.update()

    def insert_tv(self,send_info_dict):
        '''插入每个住户的信息
           循环send_info_dict：---如果value是数字(0)，那么此租户没结算房租
                              ---结算了房租：tree插入该租户
                                            创建多选按，与tree该行(item)绑定
                                            将该有效的租户信息放入self.orm
           将全选按钮设定为 打勾
           根据tree设定窗口tv_frame的高度
        :param send_info_dict:扫描excel，截图后，获取到的对应月份所有租户信息
        :return:
        '''
        #重设tree、button对应关系
        self.orm={}
        import tkinter
        for zuhu,the_month_data in send_info_dict.items():
            if type(the_month_data) !=int:  #在send_info_dict中，如果该住户没有房租信息，那么the_month_data==0
                loufang=zuhu.split('-')[0]
                fanghao=zuhu.split('-')[1]
                name=the_month_data[0]
                fangzu=the_month_data[1]
                image=the_month_data[2]
                value=[self.the_month,loufang,fanghao,name,fangzu,''] #['日期', '楼房', '房号',租客名， '房租', '发送详情']
                tv_item=self.tv.insert('', 'end', value=value,tags=('oddrow'))      #item默认状态tags
                ck_button = tkinter.Checkbutton(self.button_frame,variable=IntVar())#多选按钮
                ck_button['command']=lambda item=tv_item:self.select_button(item)  #多选按钮的回调函数对应tree里的item
                ck_button.pack()
                self.orm[tv_item]=[ck_button,zuhu,name,image] #{item:按钮、住户房号、租客名、image}

            else:#结算房租的才会添加到self.orm，否则按钮是不可按的！
                #添加tree里的item、以及对应的多选按钮，2者都显示为不可选
                value = [self.the_month, zuhu.split('-')[0], zuhu.split('-')[1], '',0, '没有此房间的房租金额']
                tv_item = self.tv.insert('', 'end', value=value, tags=('disabled'))  # item默认状态tags
                tkinter.Checkbutton(self.button_frame, state='disabled').pack()

        #每次点击插入tree，先设定全选按钮不打勾，接着打勾并且调用其函数
        self.all_buttonvar.set(0)
        self.all_button.invoke()

        #更新canvas的高度
        height = (len(self.tv.get_children()) + 1) * self.rowheight  # treeview实际高度
        self.canvas.itemconfigure(self.tv_frame, height=height) #设定窗口tv_frame的高度
        self.tk.update()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))#滚动指定的范围

    def select_all(self):
        '''回调函数---全选按钮
           作用：所有多选按钮打勾、tree所有行都改变底色(被选中)'''
        for item,[button,zuhu,name,image] in self.orm.items():
            if self.all_buttonvar.get()==1:
                button.select()
                self.tv.item(item, tags='select')
            else:
                button.deselect()
                self.tv.item(item, tags='oddrow')

    def select_button(self,item):
        '''回调函数---每个多选按钮
            作用：1.根据按钮的状态，改变对应item的底色(被选中)
                 2.根据所有按钮被选的情况，修改all_button的状态'''
        button=self.orm[item][0]
        button_value=button.getvar(button['variable'])
        if button_value=='1':
            self.tv.item(item,tags='select')
        else:
            self.tv.item(item, tags='oddrow')
        self.all_button_select()#根据所有按钮改变 全选按钮状态


    def select_tree(self,event):
        '''回调函数---点击tree的某一行
           作用：根据所点击的item改变 对应的按钮'''
        select_item=self.tv.focus()
        send_info=self.orm.get(select_item) #所选的item在orm中(结算了房租)
        if send_info:
            button = send_info[0]
            button.invoke()  #改变对应按钮的状态，而且调用其函数(改变该item颜色，也改变全选按钮)


    def all_button_select(self):
        '''根据所有按钮改变 全选按钮状态
            循环所有按钮，当有一个按钮没有被打勾时，全选按钮取消打勾'''
        for [button,zuhu,name,image] in self.orm.values():
            button_value = button.getvar(button['variable'])
            if button_value=='0':
                self.all_buttonvar.set(0)
                break
        else:
            self.all_buttonvar.set(1)
# ==================================tree与多选按钮的插入等各种操作========================###

My_Tk()

