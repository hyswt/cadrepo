from pyautocad import Autocad
import win32com.client
from tool import load_accdb,load_roughness,load_chamfer
acad=Autocad(create_if_not_exists=True)

dict_=load_accdb.func("Data_轻系列","P6",7300, "d", 30)

# dict_2=load_chamfer.func("Data_轻系列",7300, "d", 30)
# commonkeys = set(dict_.keys()) & set(dict_1.keys() & set(dict_2.keys()))
# for key in commonkeys:
#     print("相同的键：", key)

# def merge_dicts(*dict_args):
#     result = {}
#     for dictionary in dict_args:
#         # TODO 
#         result.update(dictionary)
#     return result


# dict_3=merge_dicts(dict_,dict_2)

# print(dict_3)
for text in acad.iter_objects(['Text']):
    if text.Layer == "Bstart":
        if 'rs_a' == text.TextString:
            text.TextString = ""
            gbt = load_accdb.func("Data_轻系列","P6",7300, "d", 30)
            size = gbt["r1smin"]
            upper = round(gbt["轴向"]-size, 3)
            downer = 0
            context = "\A1;{}{{\H0.5X;\S{}^{};}}".format(size, upper, downer)  # 多行文字标注上下差格式
            acad.create_mtext(text, context)
        if 'rs_r' == text.TextString:
            text.TextString = ""
            gbt = load_accdb.func("Data_轻系列","P6",7300, "d", 30)
            size = gbt["r1smin"]
            upper = round(gbt["径向"]-size,3)
            downer = 0
            context = "\A1;{}{{\H0.5X;\S{}^{};}}".format(size, upper, downer)  # 多行文字标注上下差格式
            acad.create_mtext(text, context)
        if 'r1s_r' == text.TextString:
            text.TextString = ""
            gbt = load_accdb.func("Data_轻系列","P6",7300, "d", 30)
            size = gbt["r1smin"]
            upper = round(gbt["径向"]-size,3)
            downer = 0
            context = "\A1;{}{{\H0.5X;\S{}^{};}}".format(size, upper, downer)  # 多行文字标注上下差格式
            acad.create_mtext(text, context)       
print(dict_)
for text in acad.iter_objects(['Text']):
    if text.Layer == "Bstart":
        name_=text.TextString
        # name_=text.TextOverride
        text.TextString=dict_.get(name_,f"{name_}未找到")
        pass