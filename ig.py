import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import matplotlib.pyplot as plt

class SalesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Gerenciamento de Vendas")
        self.root.geometry("1000x750")
        self.root.resizable(True, True)
        self.dados_vendas = []

        self.criar_logo()
        self.configurar_estilos()
        self.criar_formulario()
        self.criar_botoes()
        self.criar_tabela()
        self.carregar_dados_vendas()

    def criar_logo(self):
        logo_frame = ttk.Frame(self.root)
        logo_frame.pack(fill="x", padx=15, pady=10)
        ttk.Label(logo_frame, text="Sistema de Gerenciamento de Vendas", font=("Arial", 16, "bold")).pack()

    def configurar_estilos(self):
        self.root.configure(bg="#F5F5F5")
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("TLabel", font=("Arial", 11), background="#F5F5F5", foreground="#333")
        style.configure("TButton", font=("Arial", 12, "bold"), background="#4CAF50", foreground="white", padding=6)
        style.map("TButton", background=[("active", "#45A049")])
        style.configure("TEntry", font=("Arial", 11), padding=6)
        style.configure("TLabelframe", font=("Arial", 12, "bold"), background="#F5F5F5", foreground="#333", padding=10)
        style.configure("TLabelframe.Label", font=("Arial", 12, "bold"), foreground="#4CAF50")
        style.configure("Treeview", font=("Arial", 10), rowheight=30, background="white", foreground="#333")
        style.configure("Treeview.Heading", font=("Arial", 12, "bold"), background="#4CAF50", foreground="white")

    def criar_formulario(self):
        form_frame = ttk.LabelFrame(self.root, text="Dados da Venda", style="TLabelframe")
        form_frame.pack(fill="both", padx=15, pady=10, expand=True)

        canvas = tk.Canvas(form_frame, bg="#F5F5F5")
        scrollbar = ttk.Scrollbar(form_frame, orient="vertical", command=canvas.yview)
        self.form_inner_frame = ttk.Frame(canvas, style="TLabelframe")

        self.form_inner_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.form_inner_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.entradas = {}
        campos = [
            "Data da Venda", "Horário da Venda", "Produto", "SKU", "Categoria", "Método de Pagamento",
            "Região", "Cliente", "Origem do Cliente", "Status do Pedido", "Canal de Venda",
            "Quantidade", "Preço Unitário", "Desconto Aplicado (%)", "Custo Unitário", "Frete",
            "Prazo de Entrega Estimado (dias)", "Prazo de Entrega Real (dias)", "Feedback do Cliente"
        ]

        for i, campo in enumerate(campos):
            ttk.Label(self.form_inner_frame, text=campo).grid(row=i//2, column=(i%2)*2, sticky="w", padx=10, pady=5)
            entry = ttk.Entry(self.form_inner_frame, width=40)
            entry.grid(row=i//2, column=(i%2)*2 + 1, padx=10, pady=5)
            self.entradas[campo] = entry

    def criar_botoes(self):
        button_frame = ttk.Frame(self.root, style="TFrame")
        button_frame.pack(fill="x", padx=15, pady=15)

        ttk.Button(button_frame, text="Adicionar Venda", command=self.adicionar_venda).pack(side="left", padx=10, pady=10, expand=True)
        ttk.Button(button_frame, text="Exportar Dados", command=self.exportar_relatorio).pack(side="left", padx=10, pady=10, expand=True)
        ttk.Button(button_frame, text="Analisar Dados", command=self.analisar_dados).pack(side="left", padx=10, pady=10, expand=True)

    def analisar_dados(self):
        if not self.dados_vendas:
            messagebox.showwarning("Sem dados", "Não há dados para analisar.")
            return

        regioes = [venda["Região"] for venda in self.dados_vendas]
        regioes_unicas = list(set(regioes))
        vendas_por_regiao = {regiao: regioes.count(regiao) for regiao in regioes_unicas}

        regioes_custo = {regiao: [] for regiao in regioes_unicas}
        for venda in self.dados_vendas:
            try:
                custo_unitario = float(venda["Custo Unitário"])
                regioes_custo[venda["Região"]].append(custo_unitario)
            except ValueError:
                continue

        custo_medio_por_regiao = {regiao: sum(custos) / len(custos) for regiao, custos in regioes_custo.items() if custos}

        produtos = [venda["Produto"] for venda in self.dados_vendas]
        produtos_unicos = list(set(produtos))
        vendas_por_produto = {produto: produtos.count(produto) for produto in produtos_unicos}

        produtos_preco = {produto: [] for produto in produtos_unicos}
        for venda in self.dados_vendas:
            try:
                preco_unitario = float(venda["Preço Unitário"])
                produtos_preco[venda["Produto"]].append(preco_unitario)
            except ValueError:
                continue

        preco_medio_por_produto = {produto: sum(precos) / len(precos) for produto, precos in produtos_preco.items() if precos}

        fig, axs = plt.subplots(2, 2, figsize=(15, 12))

        axs[0, 0].bar(vendas_por_regiao.keys(), vendas_por_regiao.values(), color="#4CAF50")
        axs[0, 0].set_xlabel("Região")
        axs[0, 0].set_ylabel("Quantidade de Vendas")
        axs[0, 0].set_title("Quantidade de Vendas por Região")
        axs[0, 0].tick_params(axis="x", rotation=45)

        axs[0, 1].bar(custo_medio_por_regiao.keys(), custo_medio_por_regiao.values(), color="#4CAF50")
        axs[0, 1].set_xlabel("Região")
        axs[0, 1].set_ylabel("Custo Unitário Médio")
        axs[0, 1].set_title("Custo Unitário Médio por Região")
        axs[0, 1].tick_params(axis="x", rotation=45)

        axs[1, 0].bar(vendas_por_produto.keys(), vendas_por_produto.values(), color="#4CAF50")
        axs[1, 0].set_xlabel("Produto")
        axs[1, 0].set_ylabel("Quantidade de Vendas")
        axs[1, 0].set_title("Quantidade de Vendas por Produto")
        axs[1, 0].tick_params(axis="x", rotation=45)

        axs[1, 1].bar(preco_medio_por_produto.keys(), preco_medio_por_produto.values(), color="#4CAF50")
        axs[1, 1].set_xlabel("Produto")
        axs[1, 1].set_ylabel("Preço Unitário Médio")
        axs[1, 1].set_title("Preço Unitário Médio por Produto")
        axs[1, 1].tick_params(axis="x", rotation=45)

        plt.tight_layout()

        plt.savefig("analise_vendas.png")
        plt.show()

        with pd.ExcelWriter("analise_vendas.xlsx") as writer:
            pd.DataFrame(vendas_por_regiao.items(), columns=["Região", "Quantidade de Vendas"]).to_excel(writer, sheet_name="Vendas por Região", index=False)
            pd.DataFrame(custo_medio_por_regiao.items(), columns=["Região", "Custo Unitário Médio"]).to_excel(writer, sheet_name="Custo Unitário por Região", index=False)
            pd.DataFrame(vendas_por_produto.items(), columns=["Produto", "Quantidade de Vendas"]).to_excel(writer, sheet_name="Vendas por Produto", index=False)
            pd.DataFrame(preco_medio_por_produto.items(), columns=["Produto", "Preço Unitário Médio"]).to_excel(writer, sheet_name="Preço Unitário por Produto", index=False)

        with open("analise_vendas.txt", "w") as file:
            file.write("Análise de Vendas\n")
            file.write("=================\n\n")
            file.write("Quantidade de Vendas por Região:\n")
            for regiao, quantidade in vendas_por_regiao.items():
                file.write(f"{regiao}: {quantidade}\n")
            file.write("\nCusto Unitário Médio por Região:\n")
            for regiao, custo_medio in custo_medio_por_regiao.items():
                file.write(f"{regiao}: {custo_medio:.2f}\n")
            file.write("\nQuantidade de Vendas por Produto:\n")
            for produto, quantidade in vendas_por_produto.items():
                file.write(f"{produto}: {quantidade}\n")
            file.write("\nPreço Unitário Médio por Produto:\n")
            for produto, preco_medio in preco_medio_por_produto.items():
                file.write(f"{produto}: {preco_medio:.2f}\n")

        messagebox.showinfo("Análise Concluída", "A análise foi concluída e os resultados foram exportados.")

    def criar_tabela(self):
        table_frame = ttk.LabelFrame(self.root, text="Vendas Registradas", style="TLabelframe")
        table_frame.pack(fill="both", padx=15, pady=15, expand=False)

        self.tree = ttk.Treeview(table_frame, columns=list(self.entradas.keys()), show="headings")
        for col in self.entradas.keys():
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, anchor="center")
        self.tree.pack(fill=tk.BOTH, expand=True)

        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")

    def carregar_dados_vendas(self):
        try:
            if os.path.exists("relatorio_vendas_robusto.xlsx"):
                self.dados_vendas = pd.read_excel("relatorio_vendas_robusto.xlsx").to_dict(orient="records")
                for venda in self.dados_vendas:
                    self.tree.insert("", "end", values=list(venda.values()))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {e}")

    def adicionar_venda(self):
        nova_venda = {}
        for campo, entry in self.entradas.items():
            valor = entry.get()
            nova_venda[campo] = valor
        self.dados_vendas.append(nova_venda)
        self.tree.insert("", "end", values=list(nova_venda.values()))
        self.salvar_dados_vendas()
        self.limpar_formulario()
        messagebox.showinfo("Venda Adicionada", "Venda adicionada com sucesso.")

    def limpar_formulario(self):
        for campo, entry in self.entradas.items():
            entry.delete(0, tk.END)

    def exportar_relatorio(self):
        if not self.dados_vendas:
            messagebox.showwarning("Sem dados", "Não há dados para exportar.")
            return
        try:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if file_path:
                pd.DataFrame(self.dados_vendas).to_excel(file_path, index=False)
                messagebox.showinfo("Exportação Concluída", "Relatório exportado com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar relatório: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SalesApp(root)
    root.mainloop()
