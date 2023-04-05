import sys
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QLabel, QLineEdit
import pandas as pd
import webbrowser
from tkinter import *
from tkinter import filedialog
import subprocess

app = QApplication(sys.argv)

arquivo = 0

def openfile():
    filepath = filedialog.askopenfilename()
    global arquivo
    arquivo = filepath

arquivo2 = 0

def openfile2():
    filepath2 = filedialog.askopenfilename()
    global arquivo2
    arquivo2 = filepath2


janela = QWidget()
janela.resize(500, 500)
janela.setWindowTitle("Aualização de Sincronia")

button = QPushButton("Abrir Rastreabilidade.xlsx", janela)
button.setGeometry(140, 110, 200, 30)
button.clicked.connect(openfile)

button2 = QPushButton("Abrir LPN.xlsx", janela)
button2.setGeometry(140, 180, 200, 30)
button2.clicked.connect(openfile2)

texto2 = QLabel("DIGITE A ONDA", janela)
texto2.move(200, 230)
texto2.adjustSize()

texto = QLabel("SINCRONIA", janela)
texto.move(194, 25)
texto.setStyleSheet("QLabel{font-size: 18px;}")

texto3 = QLabel(
    '1. Selecione o arquivo "Rastreabilidade.xlsx" mais recente na pagina de Download', janela)
texto3.move(40, 85)
texto3.adjustSize()

texto4 = QLabel(
    '2. Selecione o arquivo "LPN.xlsx" mais recente na pagina de Download', janela)
texto4.move(40, 150)
texto4.adjustSize()

digit2 = QLineEdit("", janela)
digit2.setPlaceholderText('                     Ex. CX_SR_FL15_230303')
digit2.setGeometry(110, 260, 250, 30)


def filtro():
    pd.set_option('display.precision', 0)
    onda = digit2.text()
    tabela = pd.read_excel(arquivo, dtype=str)
    tabela["Quantidade (pç)"] = pd.to_numeric(tabela["Quantidade (pç)"])
    filtro_onda = tabela.loc[tabela['Nr. Controle 2'].str.contains(f'{onda}'), [
        "LPN Para", "De"]]

    filtro_qtd = tabela.loc[tabela['Nr. Controle 2'].str.contains(f'{onda}'), [
        "Quantidade (pç)"]]
    qtd_onda = int(filtro_qtd.sum(axis=0))

    filtro_qtd2 = tabela.loc[tabela['Nr. Controle 2'].str.contains(
        f'{onda}'), ["LPN Para", "Quantidade (pç)"]]
    filtro_qtd2.rename(columns={'LPN Para': 'LPN'}, inplace=True)
    filtro_qtd2 = pd.DataFrame(filtro_qtd2)

    qtd_LPN = filtro_qtd2.groupby(
        ["LPN"])["Quantidade (pç)"].sum().reset_index()
    qtd_LPN = pd.DataFrame(qtd_LPN)

    filtro_onda = pd.DataFrame(filtro_onda)
    filtro_onda.rename(
        columns={'LPN Para': 'LPN', "De": "ORIGEM"}, inplace=True)
    filtro_onda = filtro_onda.drop_duplicates(subset="LPN")
    tabela2 = pd.read_excel(arquivo2, dtype=str)
    tabela2["Quantidade (pç)"] = pd.to_numeric(tabela2["Quantidade (pç)"])
    tabela2 = tabela2.drop_duplicates(subset="LPN")
    col_LPN = tabela2["LPN"]
    col_LPN = pd.DataFrame(col_LPN)

    final_dados = pd.merge(tabela2, filtro_onda, on='LPN')
    final_dados = pd.DataFrame(final_dados)

    final_dados2 = final_dados["LPN"]

    tabela3 = filtro_qtd2
    tabela3["Quantidade (pç)"] = pd.to_numeric(tabela3["Quantidade (pç)"])
    filtroX = pd.merge(qtd_LPN, final_dados2, on="LPN")
    qtd_pç_final = int(filtroX["Quantidade (pç)"].sum(axis=0))
    percent_final_falta = (qtd_pç_final/qtd_onda)*100
    percent_final_realizado = (f"Realizado {100-percent_final_falta:.2f}%")
    tabela4 = {"Qt Peças Pendentes": [
        qtd_pç_final], "Realizado Sincronia": [percent_final_realizado]}
    tabela4_1 = pd.DataFrame(tabela4)

    final_dados = final_dados.drop(labels="Quantidade", axis=1)
    final_dados = final_dados.drop(labels="Quantidade (pç)", axis=1)
    final_dados = final_dados.drop(labels="Item", axis=1)
    final_dados = final_dados.drop(labels="ID Expedição", axis=1)
    final_dados = final_dados.drop(labels="Tipo", axis=1)
    final_dados = final_dados.drop(labels="LPN Pai", axis=1)
    final_dados = final_dados.drop(labels="Código do Cliente", axis=1)
    final_dados = final_dados.drop(labels="Status", axis=1)
    final_dados = final_dados.drop(labels="Tipo Estoque", axis=1)
    final_dados = final_dados.drop(labels="Ir Para", axis=1)

    final_dados2 = pd.DataFrame(final_dados)
    final_dados["Qtd Pç"] = filtroX["Quantidade (pç)"]
    final_dados2 = pd.concat([final_dados, tabela4_1])
    final_dados2 = final_dados2.fillna("")
    final_dados2 = pd.DataFrame(final_dados2)
    final_dados2.to_html('Sincronia.html')
    webbrowser.open("Sincronia.html")


def filtro2():
    pd.set_option('display.precision', 0)
    onda = digit2.text()
    tabela = pd.read_excel(arquivo, dtype=str)
    tabela["Quantidade (pç)"] = pd.to_numeric(tabela["Quantidade (pç)"])
    filtro_onda = tabela.loc[tabela['Nr. Controle 2'].str.contains(f'{onda}'), [
        "LPN Para", "De"]]

    filtro_qtd = tabela.loc[tabela['Nr. Controle 2'].str.contains(f'{onda}'), [
        "Quantidade (pç)"]]
    qtd_onda = int(filtro_qtd.sum(axis=0))

    filtro_qtd2 = tabela.loc[tabela['Nr. Controle 2'].str.contains(
        f'{onda}'), ["LPN Para", "Quantidade (pç)"]]
    filtro_qtd2.rename(columns={'LPN Para': 'LPN'}, inplace=True)
    filtro_qtd2 = pd.DataFrame(filtro_qtd2)

    qtd_LPN = filtro_qtd2.groupby(
        ["LPN"])["Quantidade (pç)"].sum().reset_index()
    qtd_LPN = pd.DataFrame(qtd_LPN)

    filtro_onda = pd.DataFrame(filtro_onda)

    filtro_onda.rename(
        columns={'LPN Para': 'LPN', "De": "ORIGEM"}, inplace=True)
    filtro_onda = filtro_onda.drop_duplicates(subset="LPN")
    tabela2 = pd.read_excel(arquivo2, dtype=str)
    tabela2["Quantidade (pç)"] = pd.to_numeric(tabela2["Quantidade (pç)"])
    tabela2 = tabela2.drop_duplicates(subset="LPN")
    col_LPN = tabela2["LPN"]
    col_LPN = pd.DataFrame(col_LPN)

    final_dados = pd.merge(tabela2, filtro_onda, on='LPN')
    final_dados = pd.DataFrame(final_dados)

    final_dados2 = final_dados["LPN"]

    tabela3 = filtro_qtd2
    tabela3["Quantidade (pç)"] = pd.to_numeric(tabela3["Quantidade (pç)"])
    filtroX = pd.merge(qtd_LPN, final_dados2, on="LPN")
    qtd_pç_final = int(filtroX["Quantidade (pç)"].sum(axis=0))
    percent_final_falta = (qtd_pç_final/qtd_onda)*100
    percent_final_realizado = (f"Realizado {100-percent_final_falta:.2f}%")
    tabela4 = {"Qt Peças Pendentes": [
        qtd_pç_final], "Realizado Sincronia": [percent_final_realizado]}
    tabela4_1 = pd.DataFrame(tabela4)

    final_dados = final_dados.drop(labels="Quantidade", axis=1)
    final_dados = final_dados.drop(labels="Quantidade (pç)", axis=1)
    final_dados = final_dados.drop(labels="Item", axis=1)
    final_dados = final_dados.drop(labels="ID Expedição", axis=1)
    final_dados = final_dados.drop(labels="Tipo", axis=1)
    final_dados = final_dados.drop(labels="LPN Pai", axis=1)
    final_dados = final_dados.drop(labels="Código do Cliente", axis=1)
    final_dados = final_dados.drop(labels="Status", axis=1)
    final_dados = final_dados.drop(labels="Tipo Estoque", axis=1)
    final_dados = final_dados.drop(labels="Ir Para", axis=1)

    final_dados2 = pd.DataFrame(final_dados)
    final_dados["Quantidade (pç)"] = filtroX["Quantidade (pç)"]
    final_dados2 = pd.concat([final_dados, tabela4_1])
    final_dados2 = final_dados2.fillna("")
    final_dados2 = pd.DataFrame(final_dados2)
    final_dados2.to_excel("Relatorio_Final.xlsx", index=False)
    subprocess.Popen("Relatorio_Final.xlsx", shell=True)


button1 = QPushButton("Abrir no Navegador", janela)
button1.setGeometry(140, 310, 200, 30)
button1.clicked.connect(filtro)

button1 = QPushButton("Abrir no Excel", janela)
button1.setGeometry(190, 350, 100, 30)
button1.clicked.connect(filtro2)


janela.show()
app.exec()
