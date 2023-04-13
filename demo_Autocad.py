from pyautocad import Autocad
import win32com.client
from tool import load_accdb,load_roughness,load_chamfer
acad=Autocad(create_if_not_exists=True)

dict_=load_accdb.func("Data_深沟球","P6",7300, "d", 30)
dict_['dd']=20
dict_['r3']=10
dict_['r8']=10
for text in acad.iter_objects(['Text']):
    if text.Layer == "Bstart":
        if 'rs_a' == text.TextString:
            text.TextString = ""
            gbt = load_accdb.func("Data_角接触","P6",7300, "d", 30)
            size = gbt["r1smin"]
            upper = round(gbt["轴向"]-size, 3)
            downer = 0
            context = "\A1;{}{{\H0.5X;\S{}^{};}}".format(size, upper, downer)  # 多行文字标注上下差格式
            acad.create_mtext(text, context)
        if 'rs_r' == text.TextString:
            text.TextString = ""
            gbt = load_accdb.func("Data_角接触","P6",7300, "d", 30)
            size = gbt["r1smin"]
            upper = round(gbt["径向"]-size,3)
            downer = 0
            context = "\A1;{}{{\H0.5X;\S{}^{};}}".format(size, upper, downer)  # 多行文字标注上下差格式
            acad.create_mtext(text, context)
        if 'r1s_r' == text.TextString:
            text.TextString = ""
            gbt = load_accdb.func("Data_角接触","P6",7300, "d", 30)
            size = gbt["r1smin"]
            upper = round(gbt["径向"]-size,3)
            downer = 0
            context = "\A1;{}{{\H0.5X;\S{}^{};}}".format(size, upper, downer)  # 多行文字标注上下差格式
            acad.create_mtext(text, context)       
print(dict_)
for text in acad.iter_objects(['Text',"MText",'Aligned','Rotated',"Linear"]):
    if text.Layer == "Bstart":
        try:
            name_=text.TextString
            text.TextString=dict_.get(name_,f"{name_}未找到")
        except :
            print(name_)
            name_=text.TextOverride
            
            text.TextOverride="\A1;{}{{\H0.5X;\S{}^{};}}".format(dict_['r3'], 0, 0)
    
        pass