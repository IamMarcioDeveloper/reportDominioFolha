import time
import pywinauto, time
from pywinauto.application import Application, timings, findwindows
from pywinauto.findwindows import ElementNotFoundError
import psutil
import pyautogui as p
from pywinauto.keyboard import send_keys

relatorios = {1:["%R_F_R", "Resumo Mensal", "FNWND3190","Resumo_Folha.pdf"],
              2: ["%R_F_E", "Extrato Mensal", "FNWND3190","Extrato_Mensal.pdf"]
              }
incrementIndex = 1
competencia = '07/2023'
erro = True
timeout = 600
emp = 3992

def close_process_by_name(process_name):
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'] == process_name:
            process.terminate()
            return True
    return False

app = Application().start(cmd_line=u'C:\Contabil\contabil.exe /Folha')
time.sleep(5)

login = Application().connect(title_re=u'Conectando .*')
login.Conectando.Edit2.type_keys("AUTOMATE")
login.Conectando.Ok.click()
time.sleep(30)

janela_principal = app.window(title_re=u'Domínio Folha .*')
time.sleep(1)

empresa = janela_principal.child_window(title_re=".*", auto_id="lblEmpresa", control_type="DominioToolbar.LabelBase")
empresa_txt = empresa.window_text()


try:
    Aviso_Vencimento = janela_principal.child_window(title="&Fechar", class_name="Button")
    Aviso_Vencimento.click()
    time.sleep(1)  # Espera um curto período de tempo após fechar a janela de aviso
except:
    pass

########################################################################################
#TROCA DE EMPRESA - EMPRESA TESTE (3992)
########################################################################################
if not empresa_txt[-4:] == emp:
    time.sleep(5)
    janela_principal.type_keys("{F8}")
    janela_principal.type_keys(f"{emp}")
    janela_principal.type_keys("{ENTER}")

try:
    Aviso_Vencimento = janela_principal.child_window(title="&Fechar", class_name="Button")
    Aviso_Vencimento.click()
    time.sleep(1)  # Espera um curto período de tempo após fechar a janela de aviso
except:
    pass

#######################################################################################################################
#CALCULO DA FOLHA
#######################################################################################################################
janela_principal.type_keys("%P_C")
time.sleep(3)
calculo = janela_principal.child_window(title="Cálculo", class_name="FNWND3190")

controle_mes_ano = janela_principal.child_window(title_re=".*", class_name="PBEDIT190")
p.hotkey(competencia)
p.hotkey("{SHIFT}")
calculo.Calcular.click()

try:
    w_handle = pywinauto.findwindows.find_window(title=u'Atenção')
    window = app.window(handle=w_handle)
    window.wait('ready')
    time.sleep(3)
    app.Aviso.Ok.click()
    time.sleep(1)
    janela_principal.type_keys("{ENTER}")
except:
    while erro:
        try:
            w_handle = pywinauto.findwindows.find_window(title_re=u'Fim .*')
            window = app.window(handle=w_handle)
            window.wait('ready')
            time.sleep(3)
            app.FimDeCálculo.No.double_click()
            erro = False
        except:
            pass
#Deseja consultar a apuração de Tributos federais?
    time.sleep(2)
    p.moveTo(x=1082, y=594)
    p.click()
#Janela de Cálculo - Opção Fechar
calculo.Fechar.click()

for x in relatorios.items():
    time.sleep(7)
    janela_principal.type_keys(relatorios[incrementIndex][0])
    time.sleep(2)
    relatorio = janela_principal.child_window(title_re=f'{relatorios[incrementIndex][1]}.*', class_name=relatorios[incrementIndex][2])
    relatorio.Ok.click()
    time.sleep(7)

    try:
        w_handle = pywinauto.findwindows.find_window(title=u'Aviso')
        window=app.window(handle=w_handle)
        window.wait('ready')
        time.sleep(3)
        app.Aviso.Ok.click()
        janela_principal.type_keys("{ESC}")
    except:
        janela_principal.type_keys("^D")
        time.sleep(3)
        w_handle = pywinauto.findwindows.find_window(title=u'Salvar em PDF')
        window=app.window(handle=w_handle)
        window.wait('ready')
        time.sleep(3)
        app.SalvarPDF.Edit.type_keys(rf"C:\Users\automate\Desktop\{emp}-{relatorios[incrementIndex][3]}")
        app.SalvarPDF.Save.double_click()
        time.sleep(5)
        try:
            process_name_to_close = 'AcroRd32.exe'
            close_process_by_name(process_name_to_close)
        except:
            pass
        janela_principal.type_keys("{ESC}{ESC}")

    incrementIndex += 1