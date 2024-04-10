import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
from datetime import datetime
import os

class KDCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("KD Calculator")

        self.tabControl = ttk.Notebook(root)

        self.tab1 = ttk.Frame(self.tabControl)
        self.tab2 = ttk.Frame(self.tabControl)

        self.tabControl.add(self.tab1, text='Rainbow Six')
        self.tabControl.add(self.tab2, text='CS2')

        self.tabControl.pack(expand=1, fill="both")

        self.mapas_r6 = ["FRONTEIRA", "CHALÉ", "CLUB HOUSE", "BANCO", "KAFE", "OREGON", "ARRANHA-CÉU", "CONSULADO", "LABORATÓRIO NIGHTHAVEN"]
        self.mapas_cs2 = ["Dust II", "Mirage", "Inferno", "Nuke", "Train", "Overpass", "Vertigo"]

        # Frame para seleção do mapa
        self.mapa_frame_r6 = ttk.LabelFrame(self.tab1, text="Selecionar Mapa")
        self.mapa_frame_r6.grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        self.mapa_var_r6 = tk.StringVar()
        self.mapa_combo_r6 = ttk.Combobox(self.mapa_frame_r6, textvariable=self.mapa_var_r6, values=self.mapas_r6, state="readonly")
        self.mapa_combo_r6.grid(row=0, column=0, padx=5, pady=5)

        self.mapa_frame_cs2 = ttk.LabelFrame(self.tab2, text="Selecionar Mapa")
        self.mapa_frame_cs2.grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        self.mapa_var_cs2 = tk.StringVar()
        self.mapa_combo_cs2 = ttk.Combobox(self.mapa_frame_cs2, textvariable=self.mapa_var_cs2, values=self.mapas_cs2, state="readonly")
        self.mapa_combo_cs2.grid(row=0, column=0, padx=5, pady=5)

        # Frame para inserção de dados
        self.dados_frame_r6 = ttk.LabelFrame(self.tab1, text="Inserir Dados")
        self.dados_frame_r6.grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
        self.dados_frame_cs2 = ttk.LabelFrame(self.tab2, text="Inserir Dados")
        self.dados_frame_cs2.grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)

        ttk.Label(self.dados_frame_r6, text="Nickname:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.nickname_entry_r6 = ttk.Entry(self.dados_frame_r6)
        self.nickname_entry_r6.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(self.dados_frame_r6, text="Data:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.data_entry_r6 = ttk.Entry(self.dados_frame_r6)
        self.data_entry_r6.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.data_entry_r6.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(self.dados_frame_r6, text="Abates:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.abates_entry_r6 = ttk.Entry(self.dados_frame_r6)
        self.abates_entry_r6.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(self.dados_frame_r6, text="Mortes:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.mortes_entry_r6 = ttk.Entry(self.dados_frame_r6)
        self.mortes_entry_r6.grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(self.dados_frame_r6, text="Assistências:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        self.assistencias_entry_r6 = ttk.Entry(self.dados_frame_r6)
        self.assistencias_entry_r6.grid(row=4, column=1, padx=5, pady=5)

        ttk.Label(self.tab1, text="Resultado:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.resultado_var_r6 = tk.StringVar()
        ttk.Radiobutton(self.tab1, text="Vitória", variable=self.resultado_var_r6, value="Vitória").grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(self.tab1, text="Derrota", variable=self.resultado_var_r6, value="Derrota").grid(row=2, column=2, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(self.tab1, text="Empate", variable=self.resultado_var_r6, value="Empate").grid(row=2, column=3, padx=5, pady=5, sticky=tk.W)

        ttk.Button(self.tab1, text="Calcular KD", command=self.calcular_kd_r6).grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)

        ttk.Label(self.dados_frame_cs2, text="Nickname:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.nickname_entry_cs2 = ttk.Entry(self.dados_frame_cs2)
        self.nickname_entry_cs2.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(self.dados_frame_cs2, text="Data:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.data_entry_cs2 = ttk.Entry(self.dados_frame_cs2)
        self.data_entry_cs2.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.data_entry_cs2.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(self.dados_frame_cs2, text="Abates:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.abates_entry_cs2 = ttk.Entry(self.dados_frame_cs2)
        self.abates_entry_cs2.grid(row=2, column=1, padx=5, pady=5)

        ttk.Label(self.dados_frame_cs2, text="Mortes:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.mortes_entry_cs2 = ttk.Entry(self.dados_frame_cs2)
        self.mortes_entry_cs2.grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(self.dados_frame_cs2, text="Assistências:").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        self.assistencias_entry_cs2 = ttk.Entry(self.dados_frame_cs2)
        self.assistencias_entry_cs2.grid(row=4, column=1, padx=5, pady=5)

        ttk.Label(self.tab2, text="Resultado:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.resultado_var_cs2 = tk.StringVar()
        ttk.Radiobutton(self.tab2, text="Vitória", variable=self.resultado_var_cs2, value="Vitória").grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(self.tab2, text="Derrota", variable=self.resultado_var_cs2, value="Derrota").grid(row=2, column=2, padx=5, pady=5, sticky=tk.W)
        ttk.Radiobutton(self.tab2, text="Empate", variable=self.resultado_var_cs2, value="Empate").grid(row=2, column=3, padx=5, pady=5, sticky=tk.W)

        ttk.Button(self.tab2, text="Calcular KD", command=self.calcular_kd_cs2).grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)

    def adicionar_cabecalho_planilha(self, nickname, jogo):
        if not os.path.exists(f"{nickname}_{jogo}.xlsx"):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Data", "Mapa", "Nickname", "Kills", "Deaths", "Assists", "Resultado", "KD"])
            wb.save(f"{nickname}_{jogo}.xlsx")

    def calcular_kd(self, nickname_entry, data_entry, abates_entry, mortes_entry, assistencias_entry, resultado_var, mapa_var, tab):
        nickname = nickname_entry.get()
        data = data_entry.get()
        abates = int(abates_entry.get())
        mortes = int(mortes_entry.get())
        assistencias = int(assistencias_entry.get())
        resultado = resultado_var.get()

        kd = round((abates + assistencias) / max(mortes, 1), 2)
        mapa_selecionado = mapa_var.get()

        jogo = tab.tab(tab.select(), "text")

        self.adicionar_cabecalho_planilha(nickname, jogo)

        wb = openpyxl.load_workbook(f"{nickname}_{jogo}.xlsx")
        ws = wb.active

        ws.append([data, mapa_selecionado, nickname, abates, mortes, assistencias, resultado, kd])

        wb.save(f"{nickname}_{jogo}.xlsx")

        messagebox.showinfo("Sucesso", "KD calculado e planilha atualizada com sucesso!")

    def calcular_kd_r6(self):
        self.calcular_kd(self.nickname_entry_r6, self.data_entry_r6, self.abates_entry_r6, self.mortes_entry_r6, self.assistencias_entry_r6, self.resultado_var_r6, self.mapa_var_r6, self.tabControl)

    def calcular_kd_cs2(self):
        self.calcular_kd(self.nickname_entry_cs2, self.data_entry_cs2, self.abates_entry_cs2, self.mortes_entry_cs2, self.assistencias_entry_cs2, self.resultado_var_cs2, self.mapa_var_cs2, self.tabControl)

def main():
    root = tk.Tk()
    app = KDCalculator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
