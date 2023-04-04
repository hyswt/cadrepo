from pyautocad import Autocad
from tool import load_accdb,load_roughness,load_chamfer
import comtypes

class CAD(Autocad):
    @property
    def app(self):
        """Returns active :class:`AutoCAD.Application`

        if :class:`Autocad` was created with :data:`create_if_not_exists=True`,
        it will create :class:`AutoCAD.Application` if there is no active one
        """
        if self._app is None:
            try:
                self._app = comtypes.client.GetActiveObject('gcad.Application', dynamic=True)
            except WindowsError:
                if not self._create_if_not_exists:
                    raise
                self._app = comtypes.client.CreateObject('gcad.Application', dynamic=True)
                self._app.Visible = self._visible
        return self._app


acad=CAD(create_if_not_exists=True)
dict_=load_accdb.func("Data_轻系列","P6", 7300,'d' ,30)
# dict_1=load_roughness.func("Data_轻系列","P6", "套圈直径", 30)
# dict_2=load_chamfer.func("Data_轻系列",7300, "d", 30)

def merge_dicts(*dict_args):
    result = {}
    for dictionary in dict_args:
        # TODO 判断keys是否有重复
        result.update(dictionary)
    return result

# dict_3=merge_dicts(dict_,dict_1,dict_2)


for text in acad.iter_objects(['Text']):
    if text.Layer == "Bstart":
        name_=text.TextString
        # name_=text.TextOverride
        text.TextString=dict_.get(name_,f"{name_}未找到")
        pass