import pypyodbc
import collections

def func(accdb_file,tb_name: str,roller_name:int,index: str, input_num: int) -> dict:
    """
    用于调用access数据库内容；\n
    使用方法
        1.将access文件放入到工作台根目录中；\n
        2.选择要使用的表格；\n
        3.输入数据进行模糊搜索。
    Args:
        accdb_file(str): 选取数据库
        tb_name (str): 所读表格的名字(P0、P2...)
        roller_name(int):轴承系列型号
        index(str): 索引标题
        input_num (int): 输入参数

    Returns:
        dict: 返回列名为key，所选行内容为value的字典
    """
    # file_path是access文件的绝对路径。
    file_path = r"./{}.accdb".format(accdb_file)
    # 链接数据库
    conn = pypyodbc.connect(u'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + file_path)
    # 创建游标
    cursor = conn.cursor()
    #NOTE 切片
    roller_name1=str(roller_name)
    roller_name2=roller_name1[-3]

    # 返回P级外圈字典
    cursor.execute('''SELECT * FROM {}级{}技术条件  WHERE {}>={}'''.format(tb_name, '外圈',index, input_num))
    columns = [column[0] for column in cursor.description]
    values = [str(value) for value in cursor.fetchone()]
    dict_ = dict(zip(columns, values))
    
    list1=['vddp9','vddp01','vddp234']
    list_1=['cir7891','cir234']

    for key in list(dict_.keys()):
        if key in list1:
            if roller_name2 in key:
                dict_['vddp']=dict_[key]
                for key in list(dict_.keys()):
                    if key in list_1:
                        if roller_name2 in key:
                            dict_['cir_dd']=dict_[key]
            

    # 返回P级内圈字典
    cursor.execute('''SELECT * FROM {}级{}技术条件  WHERE {}>={}'''.format(tb_name, '内圈',index, input_num))
    columns = [column[0] for column in cursor.description]
    values = [str(value) for value in cursor.fetchone()]
    dict1_ = dict(zip(columns, values))
    list2=['vdp9','vdp01','vdp234']
    list_2=['cir7891','cir234']
    for key in list(dict1_.keys()):
        if key in list2:
            if roller_name2 in key:
                dict1_['vdp']=dict1_[key]
                for key in list(dict1_.keys()):
                    if key in list_2:
                        if roller_name2 in key:
                            dict1_['cir_d']=dict1_[key]

    # 返回P级总图内圈字典
    cursor.execute('''SELECT * FROM {}级{}技术条件  WHERE {}>={}'''.format(tb_name, '总图内圈',index, input_num))
    columns = [column[0] for column in cursor.description]
    values = [str(value) for value in cursor.fetchone()]
    dict2_ = dict(zip(columns, values))
    list3=['zvdp89','zvdp01','zvdp234']
    for key in list(dict2_.keys()):
        if key in list3:
            if roller_name2 in key:
                dict2_['zvdp']=dict2_[key]

    # 返回P级总图外圈
    cursor.execute('''SELECT * FROM {}级{}技术条件  WHERE {}>={}'''.format(tb_name, '总图外圈',index, input_num))
    columns = [column[0] for column in cursor.description]
    values = [str(value) for value in cursor.fetchone()]
    dict3_ = dict(zip(columns, values))
    list4=['zvddp89','zvddp01','zvddp234']
    for key in list(dict3_.keys()):
        if key in list4:
            if roller_name2 in key:
                dict3_['zvddp']=dict3_[key]


    # 返回P级粗糙度字典
    cursor.execute('''SELECT * FROM {}级粗糙度表  WHERE {}>={}'''.format(tb_name,index, input_num))
    columns = [column[0] for column in cursor.description]
    values = [str(value) for value in cursor.fetchone()]
    dict4_ = dict(zip(columns, values))
    "倒角返回字典"
    
    cursor = conn.cursor()
    sql1=cursor.execute('''SELECT * FROM 直径系列{}  WHERE {}>={}'''.format(roller_name2,index, input_num))
    columns = [column[0] for column in sql1.description]
    values = [value for value in sql1.fetchone()]
    dict5_ = dict(zip(columns, values)) 
    ordered_dict = collections.OrderedDict(dict5_) 
    last_key, last_value = ordered_dict.popitem() #获取最后一列的键值
    sql_='''SELECT * FROM 倒角 where r1smin={0} and dd超过<{1} and dd到>{1}'''.format(last_value,input_num)
    cursor.execute(sql_)
    data=cursor.fetchone()
    columns1 = [column[0] for column in cursor.description]
    dict_6 = dict(zip(columns1, data))

    # 合并字典
    def merge_dicts(*dict_args):
        result = {}
        for dictionary in dict_args:
            # TODO 
            result.update(dictionary)
        return result
    dict_7=merge_dicts(dict_,dict1_,dict2_,dict3_,dict4_,dict_6)
   
    

    # #关闭游标和链接
    
    cursor.close()
    conn.close()
    return dict_7


if __name__ == "__main__":
    print(func("Data_深沟球","P6",7300, "d", 95))
    pass