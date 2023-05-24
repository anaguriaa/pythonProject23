##scrpt para entrar no sap e extrair o balancete


import win32com.client
import win32clipboard
def connectar_sap():
 SapGuiAuto = win32com.client.GetObject("SAPGUI")
 application=SapGuiAuto.GetScriptingEngine
 connection=application.Children(0)
 session=connection.Children(0)
 return session

def buscar_blancete(session):

    session.findById("wnd[0]/tbar[0]/okcd").text='f.08'
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text ="7000"
    session.findById("wnd[0]/usrtxtB_MONATE-HIGH").text=12
    session.findById("wnd[0]/usrtxtB_MONATE-HIGH").setFocus()
    session.findById("wnd[0]/usrtxtB_MONATE-HIGH").corectPosition=2
    session.findById("wnd[0]/usr").verticalScrollbatposition=1

session.findById("wnd[0]/usr").verticalScrollbatposition=2

session.findById("wnd[0]/usr").verticalScrollbatposition=3

session.findById("wnd[0]/usr").verticalScrollbatposition=4

session.findById("wnd[0]/usr").verticalScrollbatposition=5

session.findById("wnd[0]/usr").verticalScrollbatposition=6
session.findById("wnd[0]/usr[0]/mbar/menu[3]/menu[5]/menu[2]/menu[1]").select()
session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0").select()
session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0").setFocus()
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[0]/tbar[0]/btn[15]").press()
session.findById("wnd[0]/tbar[0]/btn[15]").press()
def salvar_balancet():
    win32clipboard.OpenClipboard()
    date= win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    print(data)

with open("balancete.txt","w")as f:
    for i in data:
       f.write(i)

#sap = conectar_sap()
#buscar_blancete(sap)
#salvar_balancet()


def preencher_notas(session):
    session.findById("wnd[0]/tbar[0]/okcd").text = "j1bnfe"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/txtDOCNUM-LOW").text = "10000"
    session.findById("wnd[0]/usr/ctxtDATE-LOW").text = "23.05.2023"
    session.findById("wnd[0]/usr/txtSTCD1-LOW").text = "9999999999"
    session.findById("wnd[0]/usr/ctxtMODEL-LOW").text = "55"
    session.findById("wnd[0]/usr/txtnfnum9-LOW").text = "000001234"
    session.findById("wnd[0]/usr/txtSERIE-LOW").text = "2"
    session.findById("wnd[0]/usr/txtNFYEAR-LOW").text = "22"
    session.findById("wnd[0]/usr/txtNFMONT-LOW").text = "01"
    session.findById("wnd[0]/usr/ctxtBUKRS-LOW").text = "7000"
    session.findById("wnd[0]/usr/txtBLUPA-LOW").text = ""
    session.findById("wnd[0]/usr/txtBLUPA-LOW").setFocus()
    session.findById("wnd[0]/usr/txtBLUPA-LOW").caretPosition=0
    session.findById("wnd[0]").sendVKey(0)
