摘要：最近需要将一批PDF文件中的某些数据整理到Excel中，因为文件数量接近20w+，手动更新几乎不现实，于是就提取关键词和内容动手写了个Python小工具，以实现自动完成上述目标。

---
作者：yooongchun

微信公众号： yooongchun小屋
![](yooongchun_cabin.jpg)



---
- 要求：

  - 读取PDF文件找到特定关键字，然后读取其对应的数值提取出来
  - 在Excel中查找对应关键字，然后在对应位置把上面提取出来的内容填进去

- 基本实现过程：

  - 遍历文件夹，按照特定的要求找出指定类型的PDF文件
  - 解析PDF文件
  - 提取指定内容和对应值
  - 更新数据到Excel

- 所需工具：

  - 解析PDF文件的模块：pdfminer
  - 操作Excel的模块：xlwt、xlrd、xlutils
    - 注意：要在一个已经存在的Excel中写入数据需要配合xlutils使用，即先copy一个Excel对象，在该对象中进行写入，最后删除原对象而保存copy出来的对象
  - io操作：os模块
  - 匹配PDF文件：re模块

- 代码实现：

  - 首先，把手工配置好的要求信息读入，主要包括各种文件提取规则，提取内容和文件的读写路径

    ```python
    #加载配置文件
    def loadSettingFile(KEYWORDS_Path):
        logging.info('>>>Loading setting file:%s'%os.path.basename(KEYWORDS_Path))
        PathList={}#储存路径列表
        with open(KEYWORDS_Path,'r') as fp:
            lines_kw=fp.readlines()
            for line in lines_kw:
                line=line.rstrip('\n')#删除行尾的换行符
                if re.match(r'^#',line):#注释内容，忽略
                    pass
                else:
                    Type,Path=line.split('=') #获得路径
                    PathList[Type]=Path
                    logging.info('>>>Content:\n %s'%PathList)
        logging.info('>>>Loading setting file done!')
        return PathList
    ```

  - 接着需要把刚才读入的规则按照特定的格式抽取出来

    ```python
    #提取关键词内容和值
    def extractKW(strKW):
        # 把关键词内容按照文本和数字进行分割
        logging.info('>>>Extracting key words and values from %s'%strKW)
        kw = strKW.split(';')
        key_value = {}  # 储存关键词、数据位置及列位置
        for kv in kw:
            key_value[kv.split(',')[0]] = (kv.split(',')[1],kv.split(',')[2])
            logging.info('Content:%s'%kv)
        logging.info('>>>Extracting kwywords done! ')
        return key_value
    ```

    ​

  - 使用读入的路径来初始化程序的文件操作路径

    ```python
    #初始化路径
    def InitPath(Path_List,extractKW,PDF_File_Path,Excel_Path,PDF_RULE,KeyWords,KeyWordsA117,Excel_Rule,Excel_Seri_Col,Excel_Sheet):
        folderPath=Path_List[PDF_File_Path] #PDF文件夹路径
        ExcelPath=Path_List[Excel_Path]  #Excel地址路径
        PDFRule=Path_List[PDF_RULE]    #PDF抽取规则
        kw_value=extractKW(Path_List[KeyWords])   #关键词和对应值
        kw_value_A117=extractKW(Path_List[KeyWordsA117]) #A117文件的关键词级对应值
        sheet_name=Path_List[Excel_Sheet]  #sheet名称
        xlSeriCol=Path_List[Excel_Seri_Col]  #提取序列号的列位置
        ExcelRule=Path_List[Excel_Rule]   #Excel抽取规则
        xlRule=[]  #Excel规则保存
        if not ExcelRule=='':
            for rule in ExcelRule.split(';'):
                col,con=rule.split(',')  #获得：列号 内容
                xlRule.append((int(col),con))
        return (folderPath,ExcelPath,PDFRule,xlRule,kw_value,kw_value_A117,int(xlSeriCol),sheet_name)
    ```

  - 现在到指定的目录去读取文件夹

    ```python
    #获取文件夹名称
    def loadFolder(folderPath):
        counter=0  #计数
        logging.info('>>>Loading folder from %s '%folderPath)
        folderListPath='./../folderList.txt'  #文件夹保存地址
        with open(folderListPath,'w')as f:
            folderList=os.listdir(folderPath)
            for folder in folderList:
                if not os.path.isfile(folder):
                    counter+=1
                    logging.info('>>>%s: %s'%(counter,folder))
                    f.write(os.path.join(folderPath,folder)+'\n')  #写入文件
        logging.info('>>>Done!')
        return folderListPath
    ```

  - 读取Excel，把内容加载到程序中

    ```python
    #加载Excel
    def InitExcel(excelPath):
        logging.info('>>>Loading Excel from:%s'%excelPath)
        book = xlrd.open_workbook(excelPath,formatting_info=True)  #打开一个wordbook
        copy_book= copy(book)  #拷贝一个副本
        logging.info('>>>Done!')
        return (book,copy_book)
    ```

  - 根据指定的规则来抽取Excel中的特定内容，用来之后匹配文件，找到应写入数据的对应位置

    ```python
    #抽取Excel中的序列号
    def extractExcelSeri(book,sheet_name,xlRule,xlPos):
        logging.info('>>>Extracting Excel serial from Excel Sheet:%s with xlRule:%s ...'%(sheet_name,xlRule))
        seri_data=[]#保存列数据
        sheet_ori=book.sheet_by_name(sheet_name) #切换sheet
        rows = sheet_ori.nrows #行数
        for row in range(rows-1):
            flag=True  #规则匹配标志
            for rule in xlRule:
                if (sheet_ori.cell(row,rule[0]-1).value)[0:len(rule[1])]==rule[1]:
                    pass
                else:
                    flag=False
                    break
            if flag:  #规则匹配
                seri_data.append(sheet_ori.cell(row,xlPos-1).value)
        logging.info('>>>Done!')
        return seri_data
    ```

  - 按照上面得到的文件序列来匹配文件夹名称，找到匹配的PDF文件目录

    ```python
    #使用Excel序列号匹配文件夹
    def matchFolder(xlSeri,folderListPath):
        counter=0  #计数
        logging.info('>>> Matching folder name with Excel\'s')
        matchedFolderListPath='./../matchedFolderList.txt'  #保存匹配的文件夹列表
        with open(folderListPath,'r')as f:
            lines=f.readlines()
            with open(matchedFolderListPath,'w')as ff:
                for line in lines:
                    line=line.rstrip('\n')  #去掉行尾换行符
                    line_Name=os.path.basename(line) #获取文件夹名称
                    for xlseri in xlSeri:
                        if line_Name[0:6]==xlseri[0:6]:  #序列号匹配成功
                            counter+=1
                            logging.info('>>>Matched! %s: %s'%(counter,line_Name))
                            ff.write(line+'\n')  #保存
        logging.info('>>>Done!')
        return matchedFolderListPath
    ```

  - 从前面匹配得到的PDF文件目录中抽取得到特定类型的PDF文件，抽取的规则是配置文件指定的

    ```python
    #从文件夹列表里加载指定类型的PDF文件
    def selectPDF(matchedFolderListPath,PDFRule):
        counter=0 #计数
        logging.info('>>>Loading pdf file from %s '%matchedFolderListPath)
        pdfListPath='./../pdfList.txt' #筛选出来的PDF文件列表储存位置
        with open(pdfListPath,'w')as fp:
            with open(matchedFolderListPath,'r')as f:
                folders=f.readlines()
                for folder in folders:
                    folderPath=folder.rstrip('\n')  #删除换行符
                    #遍历文件夹获取指定类型的PDF文件
                    for fpaths,dirs,fs in os.walk(folderPath):
                        for f in fs:
                            pdfName=os.path.basename(f).split('.')   #分割名称
                            if len(pdfName)>=2 and pdfName[1]=='pdf':  #判断是否属于PDF文件
                                if  re.match(PDFRule,os.path.basename(f).split('.')[0]) or 'A117' in f:  #判断是否满足PDF文件的指定规则
                                    fp.write(os.path.join(fpaths,f)+'\n')  #保存文件列表
                                    counter+=1  #计数增一
                                    logging.info('>>>%s: %s'%(counter,os.path.basename(f)))
        logging.info('>>>Selectig PDF file done!')
        return pdfListPath
    ```

  - 解析PDF文件，转换为可读取的TXT文件

    ```python
    #解析PDF文件，转为txt格式
    def parsePDF(PDF_path,TXT_path):
        logging.info('>>>Parsing pdf file:%s ...'%os.path.basename(PDF_path))
        with open(PDF_path, 'rb')as fp: # 以二进制读模式打开
            praser = PDFParser(fp)  #用文件对象来创建一个pdf文档分析器
            doc = PDFDocument() # 创建一个PDF文档
            praser.set_document(doc) # 连接分析器与文档对象
            doc.set_parser(praser)
            # 提供初始化密码
            # 如果没有密码 就创建一个空的字符串
            doc.initialize()
            # 检测文档是否提供txt转换，不提供就忽略
            if not doc.is_extractable:
                logging.info('>>>Parsing failed...')
                raise PDFTextExtractionNotAllowed
            else:
                rsrcmgr = PDFResourceManager()# 创建PDf 资源管理器 来管理共享资源
                laparams = LAParams() # 创建一个PDF设备对象
                device = PDFPageAggregator(rsrcmgr, laparams=laparams)
                interpreter = PDFPageInterpreter(rsrcmgr, device) # 创建一个PDF解释器对象

                # 循环遍历列表，每次处理一个page的内容
                for page in doc.get_pages(): # doc.get_pages() 获取page列表
                    interpreter.process_page(page)
                    layout = device.get_result() # 接受该页面的LTPage对象
                    # 这里layout是一个LTPage对象 里面存放着 这个page解析出的各种对象 一般包括LTTextBox, LTFigure, LTImage, LTTextBoxHorizontal 等等 想要获取文本就获得对象的text属性，
                    for x in layout:
                        if (isinstance(x, LTTextBoxHorizontal)):
                            with open(TXT_path, 'a',encoding='utf-8',errors='ignore') as f:
                                results = x.get_text()
                                f.write(results+'\n')
        logging.info('>>>Done!')
    ```

  - 遍历解析所有PDF文件

    ```python
    #遍历PDF列表文件完成解析
    def parseAllPDF(pdfListPath,parsePDF):
        logging.info('>>>Parsing all pdf file from pdf list:%s'%pdfListPath)
        counter=0#计数
        txtPath='./../PDF2TXT'  #保存解析好的pdf文件的路径
        if os.path.exists(txtPath): #判断目录是否存在
            pass
        else:
            os.makedirs(txtPath) #创建目录
        txtList=os.listdir(txtPath)  #加载已解析的pdf txt列表
        failed_txt_path='./failedParseList.txt'  #解析失败的文件储存位置
        with open(pdfListPath,'r') as fp:
            pdfList=fp.readlines()
            for pdfPath in pdfList:
                pdfPath=pdfPath.rstrip('\n')
                existFlag=False   #解析标志
                for file in txtList:
                    pdfName=os.path.basename(pdfPath).split('.')[0] #获取PDF文件名称
                    if file.split('.')[0]==pdfName: #判断是否已经解析过
                        logging.info('>>>This file has been parsed befores:%s/%s: %s'%(counter,len(pdfList),pdfName+'.pdf'))
                        counter+=1
                        existFlag=True
                if not existFlag:  #未曾解析过
                    counter+=1#计数
                    #生成TXT路径
                    file_Name=os.path.basename(pdfPath).split('.')[0]
                    txt_path=os.path.join('%s/%s%s'%(txtPath,file_Name,'.txt'))
                    try:
                        logging.info('>>>Parsing pdf file: %s/%s'%(counter,len(pdfList)))
                        parsePDF(pdfPath,txt_path)#解析PDF
                    except:
                        logging.info('>>>Parsing PDF:%s failed.'%os.path.basename(pdfPath))
                        with open(failed_txt_path,'a') as f: #输出错误名单
                            f.write(pdfPath+'\n')
                logging.info('>>>Done!')
        logging.info('>>>Parse all pdf file Done!')
        return txtPath
    ```

    ​

  - 从转换出来的txt文件中抽取指定内容

    ```python
    #提取TXT文件里的指定内容
    def extractContent(TXT_path,kw_value):
        logging.info('>>>Extracting content from: %s',TXT_path)
        kwv={}#储存关键字及对应值
        #读取文本内容
        with open(TXT_path,'r',encoding='utf-8',errors='ignore')as fp_tx:
            tx_lines=fp_tx.readlines()    
        if 'A117' in TXT_path:  #判断是否属于A117文件
             for con_txt in tx_lines:
                for index,item in enumerate(con_txt.split()):
                    for kw in kw_value:
                        if ' ' in kw: #判断有没有空格
                            if item==kw.split()[0] and con_txt.split()[index+1]==kw.split()[1]: #判断抽取的类型
                                if con_txt.split()[index+2]=='W': #忽略水冷类型
                                    pass
                                else:
                                    value=con_txt.split()[index+int(kw_value[kw][0])]
                                    kwv[kw]=(kw_value[kw][1],value) #返回内容，格式为：{关键字：(列号,值)}
                        else:
                            if item==kw:  #如果关键字匹配
                                if item=='PRHO':  #特殊情况
                                    value=con_txt.split()[index+int(kw_value[kw][0])]  #获得其指定位置的数据
                                else:
                                    con_txt_New=con_txt.rstrip('\n')#删除行尾的换行符
                                    value=con_txt_New.split()[int(kw_value[kw][0])]
                                kwv[kw]=(kw_value[kw][1],value) #返回内容，格式为：{关键字：(列号,值)}
        else:
            content=[]#储存内容文本
            #把文本内容按照空格分隔并存储到content中
            for con_txt in tx_lines:
                for item in con_txt.split():
                    content.append(item)

            #在文本内容中搜索关键字，找到则返回关键字及值
            for conVal,conTxt in enumerate(content):
                for kw in kw_value:#遍历关键词列表进行匹配
                    if kw==conTxt:
                        kwv[kw]=(kw_value[kw][1],content[conVal+int(kw_value[kw][0])]) #返回内容，格式为：{关键字：(列号,值)}
        logging.info('>>>Content: %s'%kwv)
        return kwv
    ```

  - 把文件内容按照匹配原则写入Excel中

    ```python
    #把指定的文本内容写入到Excel表格中
    def wtxl(kwv,kw_ori,book,copy_book,sheet_name,pdfSeri,xlPos):
        logging.info('>>>Writing data to Excel...')
        sheet_ori=book.sheet_by_name(sheet_name) #切换sheet
        rows=sheet_ori.nrows  #获得行数
        for row in range(rows-1):   #遍历行
            xlSeri=sheet_ori.cell(row,xlPos-1).value  #取得指定位置的数值
            if xlSeri==pdfSeri:  #序列号匹配成功
                sheet = copy_book.get_sheet(sheet_name) #通过sheet的名称切换
                #把内容写入到指定位置
                for kwvCon in kwv:
                    for kw in kw_ori:#遍历关键词文本
                        if kwvCon==kw and kwv[kw][1].split('.')[0].isdigit():  #匹配关键词并且关键词后面的内容为数字
                            sheet.write(row,int(kwv[kw][0])-1,kwv[kw][1])
                            logging.info('>>>Writing item:%s'%kw)
        os.remove(ExcelPath)
        copy_book.save(ExcelPath)#保存
        logging.info('>>>Done!')
    ```

  - 使用一个遍历程序把所有解析出来的PDF文件抽取内容并写入到Excel中

    ```python
    #遍历解析好的pdf文件列表提取内容并把内容写入到Excel中
    def write2Excel(Type,matchA117File,txtPath,kw_value,kw_value_A117,book,copy_book,sheet_name,xlPos,pdfListPath):
        counter=0  #计数
        logging.info('>>>Running function:write2Excel...')
        txtList=os.listdir(txtPath)
        LackOfA117ListPath='./../LackOfA117List.txt' #保存缺少A117文件列表
        with open(LackOfA117ListPath,'w')as f:
            for txt in txtList:
                counter+=1
                logging.info('>>>Dealing with PDF file: %s/%s'%(counter,len(txtList)))
                txt=txt.rstrip('\n')  #取出行尾换行符
                txtpath=txtPath+'/'+txt

                if 'Common' in txtpath:  #判断是否属于Common类型文件
                    TypeName=Type(txtpath)  #获得类型
                    if TypeName=='Direct':
                        #更新Common文件内容
                        kwv=extractContent(txtpath,kw_value)  #抽取内容
                        pdfSeri=os.path.basename(txt).split('_')[0]
                        wtxl(kwv,kw_value,book,copy_book,sheet_name,pdfSeri,xlPos)   #写入EXCEL
                        #更新Brief文件内容
                        txtpath=txtpath.replace('Common','Brief')
                        kwv=extractContent(txtpath,kw_value)  #抽取内容
                        pdfSeri=os.path.basename(txt).split('_')[0]
                        wtxl(kwv,kw_value,book,copy_book,sheet_name,pdfSeri,xlPos)   #写入EXCEL
                    else:
                        fileSeri=os.path.basename(txtpath)[0:9]  #获取文件序列号
                        a117Name=matchA117File(txtPath,fileSeri)  #获取A117文件名称
                        if not a117Name=='NO':  #该文件存在
                            a117Path=txtPath+'/'+a117Name  #获得A117文件路径
                            kwv=extractContent(a117Path,kw_value_A117)  #抽取内容
                            pdfSeri=os.path.basename(txt).split('_')[0]
                            wtxl(kwv,kw_value_A117,book,copy_book,sheet_name,pdfSeri,xlPos)   #写入EXCEL
                        else: #A117文件不存在，保存列表
                            a117Path=fileSeri+'.pdf'  #获得A117文件路径
                            f.write(a117Path+'\n')
        logging.info('>>>Done!')
    ```

  - 为了保证程序功能模块的独立，需要另外写两个小函数，分别完成获取文件类型和匹配特定类型文件的功能，这两个属于特殊情况

    ```python
    #获得指定文件的类型
    def Type(filePath):
        TypeName='NULL'
        with open(filePath,'r',encoding='utf-8',errors='ignore')as f:
            lines=f.readlines()
            for line in lines:#遍历行
                items=line.split()
                for index,item in enumerate(items):
                    if item=='Supply':  #获得类型
                        TypeName=items[index+1]
        return TypeName
      
    #匹配指定文件的A117文件
    def matchA117File(filePath,fileSeri):
        list=os.listdir(filePath)
        a117Name='NO'
        for file in list:
            if 'A117' in file:
                seri=os.path.basename(file).split('_')[0][0:9]
                if seri==fileSeri:  #匹配
                    a117Name=os.path.basename(file)
        return a117Name
    ```

    ​

  - 程序的主函数内容

    ```python

    if __name__ == '__main__':
        logging.info('>>>Program is running now...')                                     #程序开始

        ###在下面添加初始化信息
        KEYWORDS_Path='./../KEYWORDS.txt'                                                   #配置文件的路径
        PDF_File_Path='PDF_File_Path'                                                    #PDF文件夹的路径
        Excel_Path='Excel_Path'                                                          #Excel文件路径
        PDF_RULE='PDF_RULE'                                                              #PDF文件提取规则
        KeyWords='KeyWords'                                                              #关键词及值
        KeyWordsA117='KeyWordsA117'                                                      #A117文件关键词
        Excel_Rule='Excel_Rule'                                                          #Excel文件提取规则
        Excel_Seri_Col='Excel_Seri_Col'                                                  #机型匹配列位置
        Excel_Sheet='Excel_Sheet'                                                        #指定sheet名称

        ###程序运行，依次按照函数执行
        Path_List=loadSettingFile(KEYWORDS_Path)                                         #加载配置文件获取路径
                                                                                         #从配置文件内容获得相应路径
        folderPath,ExcelPath,PDFRule,xlRule,kw_value,kw_value_A117,xlSeriCol,sheet_name=InitPath(Path_List,extractKW,PDF_File_Path,Excel_Path,PDF_RULE,KeyWords,KeyWordsA117,Excel_Rule,Excel_Seri_Col,Excel_Sheet) 
        folderListPath=loadFolder(folderPath)                                            #获取文件夹名称
        book,copy_book=InitExcel(ExcelPath)                                              #初始化Excel
        xlSeri=extractExcelSeri(book,sheet_name,xlRule,xlSeriCol)                        #抽取Excel中的序列号
        matchedFolderListPath=matchFolder(xlSeri,folderListPath)                         #使用Excel序列号匹配文件夹
        pdfListPath=selectPDF(matchedFolderListPath,PDFRule)                             #从文件夹列表里加载指定类型的PDF文件
        txtPath=parseAllPDF(pdfListPath,parsePDF)                                        #遍历PDF列表文件完成解析
                                                                                         #遍历解析好的pdf文件列表提取内容并把内容写入到Excel中
        write2Excel(Type,matchA117File,txtPath,kw_value,kw_value_A117,book,copy_book,sheet_name,xlSeriCol,pdfListPath)

        logging.info('>>>Program finished!')                                             #程序完成
        input('Press any key to exit...')
    ```

- 打包为exe可执行程序

  Python程序要在没有安装Python开发包的电脑上运行的话，需要打包发布，Python提供了`pyinstaller.exe`程序来实现一键打包，首先下载安装`pyinstaller`模块，

  ```python
  pip install pyinstaller
  ```
  安装完成后搜索找到`pyinstaller.exe` 复制到你想要打包的文件的位置，也就是你的`.py` 文件的位置，然后使用命令行执行：

  ```python
  cd 你的上述文件放置位置
  pyinstaller.exe 你的.py文件名称
  ```

  比如我的`pyinstaller.exe` 放在了`C:/Users/fanyu/desktop/Python` 路径下，同时里面还有一个`TEST.py` 的文件我想要打包成`exe` 程序，那么我的运行命令就是：

  ```python
  cd C:/users/fanyu/desktop/Python
  pyinstaller.exe TEST.py
  ```

  现在如果一切正常的话程序就会运行在当前目录下生成`dist` 、`build` 、`TEST.spec` 、`_pycache_` 的四个文件，需要的运行程序在`dist` 目录下，里面除了`exe` 程序外会有许多文件，那是程序运行需要的支持文件。

  当然，`pyinstaller.exe` 还提供了更丰富的打包功能，比如加入自己的程序图标，程序运行时不显示命令行窗口等，这个就自己探索了！

- 程序的使用说明文档:`使用说明.txt`

  ```wiki
  这个工具用来实现从指定的文件夹读取文件，抽取特定的数据写入到指定的Excel文件中的功能

  1.如何使用？
    你需要在KEYWORDS.txt文件中填写路径以及规则，然后运行TOOL文件夹下的TOOL.exe即可

  2.程序运行逻辑？
     a-->打开Excel文件按照指定的规则取得值
  		
     b-->遍历指定的PDF文件夹，将其名称与Excel中得到的进行匹配，若匹配成功，则保存该文件或文件夹的路径到folderList.txt中

     c-->遍历上述文件夹内的按照指定规则获得的所有PDF文件并保存到pdfList.txt中

     d-->解析上面获得的PDF文件并保存到PDF2Txt文件夹中
     
     e-->按照指定的规则抽取d步骤获得的TXT文件中的内容

     f-->把e步骤获得的内容写入到对应的Excel位置

     g-->程序完成

  3.如何配置规则？
    所有规则需要在运行程序前在KEYWORDS.txt文件中配置，包括：
    
    a-->PDF文件夹所在路径
    b-->Excel文件所在路径
    c-->PDF文件的提取规则
    d-->Excel抽取规则
    e-->Excel写入规则
    f-->抽取数据位置及写入Excel位置规则

  4.如何获得帮助？
   
    联系：fanyu Email:1729465178@qq.com   QQ:1729465178
  ```

- 程序代码和exe程序下载：https://github.com/yooongchun/Python_pdf2Excel