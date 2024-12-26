import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import matplotlib.pyplot as plt
import pandas as pd
import matplotlib.pyplot as plt
from tkinter import messagebox
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

import os
import pandas as pd
import matplotlib.pyplot as plt
from tkinter import messagebox
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader



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

        df = pd.DataFrame(self.dados_vendas)

        required_columns = {'Data da Venda', 'Horário da Venda', 'Produto', 'SKU', 'Categoria',
                            'Método de Pagamento', 'Região', 'Cliente', 'Origem do Cliente',
                            'Status do Pedido', 'Canal de Venda', 'Quantidade', 'Preço Unitário',
                            'Desconto Aplicado (%)', 'Custo Unitário', 'Total da Venda',
                            'Lucro Bruto', 'Frete', 'Prazo de Entrega Estimado (dias)',
                            'Prazo de Entrega Real (dias)', 'Feedback do Cliente'}

        if not required_columns.issubset(df.columns):
            messagebox.showerror("Erro", "Faltam colunas necessárias nos dados para análise.")
            return

        if not os.path.exists("img"):
            os.makedirs("img")

        vendas_por_regiao = df['Região'].value_counts()
        vendas_por_produto = df['Produto'].value_counts()
        preco_medio_por_produto = df.groupby('Produto')['Preço Unitário'].mean()

        cancelados = df[df['Status do Pedido'] == 'Cancelado']
        bins = [0, 10, 15, 20, 30, 50, 100]
        labels = ['0-10', '10-15', '15-20', '20-30', '30-50', '50-100']
        cancelados['Faixa de Frete'] = pd.cut(cancelados['Frete'], bins=bins, labels=labels)

        cancelamentos_por_faixa = cancelados['Faixa de Frete'].value_counts().sort_index()
        frete_medio_por_faixa = cancelados.groupby('Faixa de Frete')['Frete'].mean()

        regiao_com_mais_cancelamentos = df[df['Status do Pedido'] == 'Cancelado']['Região'].value_counts()
        regiao_com_mais_clientes_insatisfeitos = df[df['Feedback do Cliente'] == 'Muito Insatisfeito']['Região'].value_counts()
        produtos_com_mais_insatisfacoes = df[df['Feedback do Cliente'] == 'Insatisfeito']['Produto'].value_counts()

        frete_medio_por_produto = df.groupby('Produto')['Frete'].mean()
        correlacao_frete_vendas = frete_medio_por_produto.corr(vendas_por_produto)

        produtos_bem_vendidos_por_regiao = df.groupby(['Região', 'Produto'])['Quantidade'].sum().unstack().fillna(0)

        def salvar_grafico(fig, filename):
            fig.tight_layout()
            fig.savefig(filename)
            plt.close(fig)

        fig, ax = plt.subplots(figsize=(9, 5))
        ax.bar(vendas_por_regiao.index, vendas_por_regiao.values, color="#4CAF50")
        ax.set_title("Quantidade de Vendas por Região")
        ax.set_xlabel("Região")
        ax.set_ylabel("Quantidade de Vendas")
        ax.tick_params(axis="x", rotation=45)
        salvar_grafico(fig, "img/vendas_por_regiao.png")

        fig, ax = plt.subplots(figsize=(9, 5))
        ax.bar(vendas_por_produto.index, vendas_por_produto.values, color="#4CAF50")
        ax.set_title("Quantidade de Vendas por Produto")
        ax.set_xlabel("Produto")
        ax.set_ylabel("Quantidade de Vendas")
        ax.tick_params(axis="x", rotation=45)
        salvar_grafico(fig, "img/vendas_por_produto.png")

        fig, ax = plt.subplots(figsize=(9, 5))
        ax.bar(cancelamentos_por_faixa.index.astype(str), cancelamentos_por_faixa.values, color="#4CAF50")
        ax.set_title("Cancelamentos por Faixa de Frete")
        ax.set_xlabel("Faixa de Frete")
        ax.set_ylabel("Cancelamentos")
        ax.tick_params(axis="x", rotation=45)
        salvar_grafico(fig, "img/cancelamentos_por_faixa.png")

        fig, ax = plt.subplots(figsize=(9, 5))
        ax.bar(frete_medio_por_produto.index, frete_medio_por_produto.values, color="#4CAF50")
        ax.set_title("Frete Médio por Produto")
        ax.set_xlabel("Produto")
        ax.set_ylabel("Frete Médio")
        ax.tick_params(axis="x", rotation=45)
        salvar_grafico(fig, "img/frete_medio_por_produto.png")

        fig, ax = plt.subplots(figsize=(9, 5))
        produtos_bem_vendidos_por_regiao.plot(kind='bar', stacked=True, ax=ax, colormap='viridis')
        ax.set_title("Produtos Bem Vendidos por Região")
        ax.set_xlabel("Região")
        ax.set_ylabel("Quantidade Vendida")
        ax.legend(title="Produto", bbox_to_anchor=(1.05, 1), loc='upper left')
        salvar_grafico(fig, "img/produtos_bem_vendidos_por_regiao.png")

        conclusoes = []

        conclusoes.append("Análise de Cancelamentos por Faixa de Frete:")
        for faixa, cancelamentos in cancelamentos_por_faixa.items():
            conclusoes.append(f"- Faixa {faixa}: {cancelamentos} cancelamentos")

        conclusoes.append("\nFrete Médio por Faixa de Frete:")
        for faixa, frete in frete_medio_por_faixa.items():
            conclusoes.append(f"- Faixa {faixa}: R$ {frete:.2f}")

        conclusoes.append("\nRegião com o Maior Número de Cancelamentos:")
        for regiao, cancelamentos in regiao_com_mais_cancelamentos.items():
            conclusoes.append(f"- Região {regiao}: {cancelamentos} cancelamentos")

        conclusoes.append("\nRegião com o Maior Número de Clientes Insatisfeitos:")
        for regiao, insatisfeitos in regiao_com_mais_clientes_insatisfeitos.items():
            conclusoes.append(f"- Região {regiao}: {insatisfeitos} clientes insatisfeitos")

        conclusoes.append("\nProduto com Maior Número de Insatisfações:")
        for produto, insatisfeitos in produtos_com_mais_insatisfacoes.items():
            conclusoes.append(f"- Produto {produto}: {insatisfeitos} insatisfações")

        conclusoes.append("\nAnálise de Correlação entre Frete e Vendas:")
        conclusoes.append(f"A correlação entre frete médio e quantidade de vendas por produto é: {correlacao_frete_vendas:.2f}")

        conclusoes_texto = "\n".join(conclusoes)
        print(conclusoes_texto)

        pdf_filename = "analise_vendas.pdf"
        c = canvas.Canvas(pdf_filename, pagesize=letter)
        width, height = letter

        c.setFont("Helvetica", 12)
        c.drawString(30, height - 30, "Análise de Dados de Vendas")
        c.setFont("Helvetica", 10)

        y_position = height - 60
        for linha in conclusoes:
            if linha.startswith("\n"):
                y_position -= 20
                linha = linha.strip()
            c.drawString(30, y_position, linha)
            y_position -= 15
            if y_position < 40:
                c.showPage()
                c.setFont("Helvetica", 10)
                y_position = height - 40

        c.showPage()
        c.drawImage(ImageReader("img/vendas_por_regiao.png"), 30, 30, width - 60, height - 60)
        c.showPage()
        c.drawImage(ImageReader("img/vendas_por_produto.png"), 30, 30, width - 60, height - 60)
        c.showPage()
        c.drawImage(ImageReader("img/cancelamentos_por_faixa.png"), 30, 30, width - 60, height - 60)
        c.showPage()
        c.drawImage(ImageReader("img/frete_medio_por_produto.png"), 30, 30, width - 60, height - 60)
        c.showPage()
        c.drawImage(ImageReader("img/produtos_bem_vendidos_por_regiao.png"), 30, 30, width - 60, height - 60)

        c.save()

        messagebox.showinfo("Análise Concluída", f"A análise foi concluída com sucesso e os gráficos foram exibidos.\n\n{conclusoes_texto}\n\nO relatório foi salvo como {pdf_filename}")
                
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
