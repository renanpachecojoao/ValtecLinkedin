"""
Valtec LinkedIn Analyzer — Versão Profissional
Análise completa de Conteúdo + Visitantes + Seguidores
Dependências: pip install groq matplotlib pandas openpyxl xlrd
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import json
import re
import os
import pandas as pd
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from groq import Groq

# ── Paleta de cores ────────────────────────────────────────────────────────────
COR_AZUL      = "#0A66C2"
COR_AZUL_ESC  = "#0A3D6B"
COR_AZUL_CLR  = "#EBF3FB"
COR_VERDE     = "#1D9E75"
COR_VERDE_CLR = "#E1F5EE"
COR_AMARELO   = "#E6A817"
COR_AMAR_CLR  = "#FEF9EC"
COR_VERM      = "#C0392B"
COR_VERM_CLR  = "#FCEBEB"
COR_CINZA     = "#F0F2F5"
COR_CINZA2    = "#E4E6EA"
COR_BORDA     = "#D0D7DE"
COR_TEXTO     = "#1A1A2E"
COR_MUTED     = "#6B7280"
COR_BRANCO    = "#FFFFFF"
COR_HEADER    = "#0A3D6B"
COR_ROXO      = "#7C3AED"
COR_ROXO_CLR  = "#F3F0FF"

FONTES = {
    "titulo":   ("Helvetica", 13, "bold"),
    "subtit":   ("Helvetica", 11, "bold"),
    "body":     ("Helvetica", 10),
    "small":    ("Helvetica", 9),
    "metric":   ("Helvetica", 18, "bold"),
    "score":    ("Helvetica", 36, "bold"),
    "header":   ("Helvetica", 16, "bold"),
}

# ── IA com Groq ────────────────────────────────────────────────────────────────
def analisar_com_ia(summary: str, api_key: str) -> dict:
    client = Groq(api_key=api_key.strip())

    contexto_empresa = """
SOBRE A EMPRESA (use isso para calibrar insights, calendário editorial e recomendações — NÃO genéricos de "engenharia industrial", mas específicos ao negócio real abaixo):

Nome: Valtec Instalação
Localização: Mauá/SP — atuação em todo o Brasil

O que a empresa faz:
Desenvolve e implementa soluções de engenharia inovadoras e personalizadas, com excelência, segurança, eficiência e sustentabilidade. Promove uso responsável de recursos humanos, ambientais e financeiros, com foco em preservação ambiental.

Setores de atuação: construção civil e plantas industriais — mineração, papel e celulose, automobilística, fertilizantes, química e petrolífera.

Serviços oferecidos (use como base para sugestões de temas de conteúdo):
- Montagens mecânicas e eletromecânicas
- Instrumentação
- Instalações elétricas e hidráulicas
- Sistemas de combate a incêndio
- Sistemas eletrônicos diversos

Diferenciais: atuação arrojada, foco em qualidade, cumprimento de prazos e satisfação do cliente, melhoria contínua de processos, soluções inovadoras.

Missão: fornecer soluções técnicas inovadoras e personalizadas que superam expectativas dos clientes, com eficiência, segurança e qualidade, além de focar no desenvolvimento da comunidade e bem-estar dos colaboradores.

Visão: ser reconhecida regional e nacionalmente como uma das principais empresas de Engenharia em Construção Civil, Montagem Eletromecânica e Instalações, destacando-se por integridade, inovação, sustentabilidade e excelência operacional.

Valores: Sustentabilidade, Integridade, Segurança, Comprometimento, Inovação.

Público-alvo no LinkedIn: engenheiros, gerentes e diretores de obras/plantas industriais nos setores acima, profissionais de EHS (segurança/meio ambiente), compradores e decisores técnicos de grandes indústrias.
"""

    prompt = f"""Você é especialista sênior em LinkedIn Analytics para empresas B2B de engenharia industrial no Brasil.
Analise os dados completos da página da Valtec Instalação no LinkedIn (~5 mil seguidores).

{contexto_empresa}

IMPORTANTE: Use APENAS dados reais encontrados nos arquivos para métricas e números. Se um dado não existir, indique "N/D".
Para insights, padrões de sucesso, calendário editorial e recomendações, baseie-se no contexto real da empresa acima (setores, serviços, missão, valores) — não generalize para "engenharia industrial" de forma vaga. Sugira temas de conteúdo concretos ligados aos serviços (montagem eletromecânica, instrumentação, elétrica/hidráulica, combate a incêndio) e aos setores atendidos (mineração, papel e celulose, automotivo, fertilizantes, química, petróleo).

Retorne APENAS JSON válido, sem markdown, sem texto extra:

{{
  "score_geral": 72,
  "score_interpretacao": "Sua página está acima da média B2B industrial",
  "resumo": "2 frases objetivas resumindo a situação geral da página com base nos dados reais",
  "periodo": "Ex: 01/05/2025 a 31/05/2025",
  "metrics": [
    {{"label": "Total de Impressões",    "value": "...", "variacao": "+12%", "status": "good|warn|bad"}},
    {{"label": "Taxa de Engajamento",    "value": "...", "variacao": "+2%",  "status": "good|warn|bad"}},
    {{"label": "Novos Seguidores",       "value": "...", "variacao": "+5%",  "status": "good|warn|bad"}},
    {{"label": "Visitantes Únicos",      "value": "...", "variacao": "0%",   "status": "good|warn|bad"}},
    {{"label": "Cliques no Link",        "value": "...", "variacao": "-3%",  "status": "good|warn|bad"}},
    {{"label": "Compartilhamentos",      "value": "...", "variacao": "+1%",  "status": "good|warn|bad"}},
    {{"label": "Alcance Médio por Post", "value": "...", "variacao": "+8%",  "status": "good|warn|bad"}},
    {{"label": "Posts Publicados",       "value": "...", "variacao": "0%",   "status": "good|warn|bad"}}
  ],
  "score_por_pilar": [
    {{"pilar": "Conteúdo",    "score": 0, "descricao": "..."}},
    {{"pilar": "Alcance",     "score": 0, "descricao": "..."}},
    {{"pilar": "Engajamento", "score": 0, "descricao": "..."}},
    {{"pilar": "Crescimento", "score": 0, "descricao": "..."}},
    {{"pilar": "Audiência",   "score": 0, "descricao": "..."}}
  ],
  "melhores_dias": [
    {{"dia": "Seg", "score": 0}},
    {{"dia": "Ter", "score": 0}},
    {{"dia": "Qua", "score": 0}},
    {{"dia": "Qui", "score": 0}},
    {{"dia": "Sex", "score": 0}},
    {{"dia": "Sáb", "score": 0}},
    {{"dia": "Dom", "score": 0}}
  ],
  "melhores_horarios": [
    {{"hora": "6h-8h",   "score": 0}},
    {{"hora": "8h-10h",  "score": 0}},
    {{"hora": "10h-12h", "score": 0}},
    {{"hora": "12h-14h", "score": 0}},
    {{"hora": "14h-16h", "score": 0}},
    {{"hora": "16h-18h", "score": 0}},
    {{"hora": "18h-20h", "score": 0}},
    {{"hora": "20h-22h", "score": 0}}
  ],
  "tipos_conteudo": [
    {{"tipo": "Vídeo",   "alcance": 0, "engajamento": 0.0, "frequencia": 0}},
    {{"tipo": "Imagem",  "alcance": 0, "engajamento": 0.0, "frequencia": 0}},
    {{"tipo": "Texto",   "alcance": 0, "engajamento": 0.0, "frequencia": 0}},
    {{"tipo": "Artigo",  "alcance": 0, "engajamento": 0.0, "frequencia": 0}}
  ],
  "audiencia": {{
    "cargos_top": [
      {{"cargo": "Engenheiro", "percentual": 0}},
      {{"cargo": "Gerente",    "percentual": 0}},
      {{"cargo": "Diretor",    "percentual": 0}},
      {{"cargo": "Técnico",    "percentual": 0}},
      {{"cargo": "Outros",     "percentual": 0}}
    ],
    "setores_top": [
      {{"setor": "Manufatura",    "percentual": 0}},
      {{"setor": "Construção",    "percentual": 0}},
      {{"setor": "Energia",       "percentual": 0}},
      {{"setor": "Outros",        "percentual": 0}}
    ],
    "origem_seguidores": [
      {{"origem": "Busca orgânica", "percentual": 0}},
      {{"origem": "Sugestão LI",    "percentual": 0}},
      {{"origem": "Post",           "percentual": 0}},
      {{"origem": "Outros",         "percentual": 0}}
    ]
  }},
  "top_posts": [
    {{
      "titulo": "Título/tema do post",
      "tipo": "Imagem|Vídeo|Texto|Artigo",
      "impressoes": 0,
      "engajamento": 0.0,
      "cliques": 0,
      "status": "top|low",
      "padrao": "O que fez esse post funcionar ou falhar",
      "motivo": "Análise mais detalhada"
    }}
  ],
  "padroes_sucesso": [
    {{"padrao": "...", "impacto": "alto|medio|baixo", "exemplo": "..."}}
  ],
  "insights": [
    {{"tipo": "erro", "titulo": "...", "texto": "...", "impacto": "alto|medio"}},
    {{"tipo": "erro", "titulo": "...", "texto": "...", "impacto": "alto|medio"}},
    {{"tipo": "ok",   "titulo": "...", "texto": "...", "impacto": "alto|medio"}},
    {{"tipo": "ok",   "titulo": "...", "texto": "...", "impacto": "alto|medio"}},
    {{"tipo": "dica", "titulo": "...", "texto": "...", "impacto": "alto|medio"}},
    {{"tipo": "dica", "titulo": "...", "texto": "...", "impacto": "alto|medio"}}
  ],
  "calendario_sugerido": [
    {{
      "semana": "Semana 1",
      "posts": [
        {{"dia": "Terça", "horario": "10h", "tipo": "Imagem", "tema": "Case de obra finalizada", "justificativa": "Posts de resultado têm 2x mais alcance"}},
        {{"dia": "Quinta", "horario": "08h", "tipo": "Vídeo", "tema": "Processo de instalação elétrica", "justificativa": "Vídeos geram 3x mais engajamento"}}
      ]
    }},
    {{
      "semana": "Semana 2",
      "posts": [
        {{"dia": "Quarta", "horario": "10h", "tipo": "Texto", "tema": "Dica técnica de segurança", "justificativa": "Conteúdo educativo atrai engenheiros"}},
        {{"dia": "Sexta",  "horario": "09h", "tipo": "Imagem","tema": "Equipe em campo", "justificativa": "Humaniza a marca e gera engajamento"}}
      ]
    }}
  ],
  "metas_30dias": [
    {{"meta": "Impressões", "atual": "...", "alvo": "...", "acao": "..."}},
    {{"meta": "Seguidores", "atual": "...", "alvo": "...", "acao": "..."}},
    {{"meta": "Engajamento","atual": "...", "alvo": "...", "acao": "..."}}
  ],
  "recomendacoes": [
    {{"titulo": "...", "descricao": "Recomendação detalhada e prática para engenharia industrial", "prioridade": "alta|media|baixa", "prazo": "Imediato|30 dias|60 dias"}}
  ]
}}

Scores de 0-100. Máximo 4 top_posts (2 melhores + 2 piores). Máximo 3 padroes_sucesso. Máximo 5 recomendacoes.
Use dados reais dos arquivos quando disponíveis. Para campos sem dados, use melhores práticas B2B engenharia Brasil.

DADOS DO LINKEDIN (Conteúdo + Visitantes + Seguidores):
{summary}
"""

    resp = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=3000,
    )
    raw = resp.choices[0].message.content.strip()
    raw = re.sub(r"```json|```", "", raw).strip()
    return json.loads(raw)


def preparar_resumo(frames: dict) -> str:
    linhas = []
    for nome, df in frames.items():
        linhas.append(f"=== {nome} ({len(df)} linhas) ===")
        linhas.append(f"Colunas: {', '.join(df.columns.astype(str))}")
        linhas.append(df.head(50).to_string(index=False))
    return "\n".join(linhas)[:12000]


# ── App Principal ──────────────────────────────────────────────────────────────
class ValtecApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Valtec LinkedIn Analyzer — Profissional")
        self.geometry("1280x820")
        self.minsize(1000, 640)
        self.configure(bg=COR_CINZA)

        self.dados = {}          # acumula todos os arquivos importados
        self.arquivos_nome = []  # nomes dos arquivos carregados
        self.api_key = tk.StringVar()
        self._build_ui()

    # ── Construção da UI ───────────────────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=COR_HEADER, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  ⚡ Valtec LinkedIn Analyzer",
                 font=("Helvetica", 17, "bold"), bg=COR_HEADER, fg=COR_BRANCO).pack(side="left", padx=16)
        tk.Label(hdr, text="Análise Profissional  •  Groq + Llama 3.3",
                 font=("Helvetica", 10), bg=COR_HEADER, fg="#90CAF9").pack(side="right", padx=16)

        # Layout: lateral + principal
        corpo = tk.Frame(self, bg=COR_CINZA)
        corpo.pack(fill="both", expand=True)

        # ── Painel lateral ─────────────────────────────────────────────────
        lateral = tk.Frame(corpo, bg=COR_BRANCO, width=290,
                           highlightbackground=COR_BORDA, highlightthickness=1)
        lateral.pack(side="left", fill="y", padx=(10, 0), pady=10)
        lateral.pack_propagate(False)

        # Scroll no lateral
        lat_canvas = tk.Canvas(lateral, bg=COR_BRANCO, highlightthickness=0)
        lat_scroll  = ttk.Scrollbar(lateral, orient="vertical", command=lat_canvas.yview)
        lat_canvas.configure(yscrollcommand=lat_scroll.set)
        lat_scroll.pack(side="right", fill="y")
        lat_canvas.pack(side="left", fill="both", expand=True)
        lat_inner = tk.Frame(lat_canvas, bg=COR_BRANCO)
        lat_win = lat_canvas.create_window((0, 0), window=lat_inner, anchor="nw")
        def _lat_resize(e): lat_canvas.itemconfig(lat_win, width=e.width)
        def _lat_frame(e):  lat_canvas.configure(scrollregion=lat_canvas.bbox("all"))
        lat_canvas.bind("<Configure>", _lat_resize)
        lat_inner.bind("<Configure>", _lat_frame)

        # API Key
        self._secao(lat_inner, "🔑 Groq API Key")
        self.entry_key = ttk.Entry(lat_inner, textvariable=self.api_key, show="*", width=32)
        self.entry_key.pack(padx=12, pady=(0, 2), fill="x")
        lnk = tk.Label(lat_inner, text="→ Obter chave grátis em console.groq.com",
                 fg=COR_AZUL, bg=COR_BRANCO, font=("Helvetica", 8), cursor="hand2")
        lnk.pack(padx=12, anchor="w")

        self._sep(lat_inner)

        # Importar arquivos (acumula vários)
        self._secao(lat_inner, "📂 Importar relatórios do LinkedIn")
        tk.Label(lat_inner,
                 text="Exporte 3 arquivos separados:\n• Analytics → Conteúdo → Exportar\n• Analytics → Visitantes → Exportar\n• Analytics → Seguidores → Exportar\n\nDica: você pode selecionar os 3\narquivos de uma vez (Ctrl+clique).",
                 bg=COR_BRANCO, fg=COR_MUTED, font=("Helvetica", 9), justify="left").pack(padx=12, anchor="w")

        self._botao(lat_inner, "＋  Adicionar arquivo(s) (.xlsx/.xls/.csv)", self._importar, COR_AZUL)

        # Lista de arquivos carregados
        self.frame_arquivos = tk.Frame(lat_inner, bg=COR_BRANCO)
        self.frame_arquivos.pack(fill="x", padx=12, pady=(6, 0))
        self._atualizar_lista_arquivos()

        self._sep(lat_inner)

        # Dados manuais
        self._secao(lat_inner, "📋 Ou cole dados manualmente")
        tk.Label(lat_inner, text="Cole qualquer texto copiado\ndo LinkedIn Analytics:",
                 bg=COR_BRANCO, fg=COR_MUTED, font=("Helvetica", 9), justify="left").pack(padx=12, anchor="w")
        self.texto_manual = scrolledtext.ScrolledText(
            lat_inner, height=7, width=30, font=("Courier", 8),
            bg="#F9FAFB", relief="flat",
            highlightbackground=COR_BORDA, highlightthickness=1)
        self.texto_manual.pack(padx=12, pady=(4, 0), fill="x")

        self._sep(lat_inner)

        # Botão analisar
        self.btn_analisar = self._botao(lat_inner, "✨  Gerar Análise Completa", self._analisar, COR_VERDE)

        # Botão limpar
        self._botao(lat_inner, "🗑  Limpar tudo", self._limpar, COR_MUTED)

        # Status
        self.label_status = tk.Label(lat_inner, text="", bg=COR_BRANCO,
                                     fg=COR_MUTED, font=("Helvetica", 9), wraplength=250)
        self.label_status.pack(padx=12, pady=6, anchor="w")

        tk.Label(lat_inner, text="", bg=COR_BRANCO).pack(pady=10)

        # ── Área principal ─────────────────────────────────────────────────
        self.area_principal = tk.Frame(corpo, bg=COR_CINZA)
        self.area_principal.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        self._tela_inicial()

    # ── Helpers de UI ──────────────────────────────────────────────────────────
    def _secao(self, parent, titulo):
        tk.Label(parent, text=titulo, bg=COR_BRANCO, fg=COR_TEXTO,
                 font=FONTES["subtit"]).pack(padx=12, pady=(12, 4), anchor="w")

    def _sep(self, parent):
        ttk.Separator(parent, orient="horizontal").pack(fill="x", padx=12, pady=8)

    def _botao(self, parent, texto, cmd, cor):
        btn = tk.Button(parent, text=texto, command=cmd,
                        bg=cor, fg=COR_BRANCO, activebackground=COR_AZUL_ESC,
                        activeforeground=COR_BRANCO, relief="flat", cursor="hand2",
                        font=("Helvetica", 10, "bold"), pady=9)
        btn.pack(padx=12, pady=(4, 0), fill="x")
        return btn

    def _card(self, parent, bg=COR_BRANCO, borda=COR_BORDA, bw=1, pad=(8, 8)):
        frm = tk.Frame(parent, bg=bg,
                       highlightbackground=borda, highlightthickness=bw)
        frm.pack(fill="x", padx=4, pady=3)
        return frm

    def _titulo_secao(self, parent, texto):
        frm = tk.Frame(parent, bg=COR_CINZA)
        frm.pack(fill="x", padx=4, pady=(14, 5))
        tk.Label(frm, text=texto, font=FONTES["subtit"],
                 bg=COR_CINZA, fg=COR_TEXTO).pack(side="left")
        tk.Frame(frm, bg=COR_BORDA, height=1).pack(side="left", fill="x",
                                                     expand=True, padx=(8, 0), pady=6)

    def _atualizar_lista_arquivos(self):
        for w in self.frame_arquivos.winfo_children():
            w.destroy()
        if not self.arquivos_nome:
            tk.Label(self.frame_arquivos, text="Nenhum arquivo carregado ainda.",
                     bg=COR_BRANCO, fg=COR_MUTED, font=FONTES["small"]).pack(anchor="w")
        else:
            for i, nome in enumerate(self.arquivos_nome):
                row = tk.Frame(self.frame_arquivos, bg=COR_BRANCO)
                row.pack(fill="x", pady=1)
                tk.Label(row, text=f"✓ {nome}", bg=COR_BRANCO,
                         fg=COR_VERDE, font=FONTES["small"],
                         wraplength=210, justify="left").pack(side="left")

    # ── Tela inicial ───────────────────────────────────────────────────────────
    def _tela_inicial(self):
        for w in self.area_principal.winfo_children():
            w.destroy()
        frm = tk.Frame(self.area_principal, bg=COR_CINZA)
        frm.place(relx=0.5, rely=0.45, anchor="center")

        # Ícone grande
        tk.Label(frm, text="📊", font=("Helvetica", 64), bg=COR_CINZA).pack()
        tk.Label(frm, text="Análise Profissional de LinkedIn",
                 font=("Helvetica", 18, "bold"), bg=COR_CINZA, fg=COR_TEXTO).pack(pady=(8, 4))
        tk.Label(frm, text="Importe os 3 relatórios do LinkedIn Analytics\ne clique em  ✨ Gerar Análise Completa",
                 font=("Helvetica", 12), bg=COR_CINZA, fg=COR_MUTED, justify="center").pack()

        # Cards de instruções
        cards_frm = tk.Frame(frm, bg=COR_CINZA)
        cards_frm.pack(pady=20)
        passos = [
            ("1", "📂", "Importe os relatórios", "Conteúdo, Visitantes\ne Seguidores"),
            ("2", "🔑", "Insira a API Key", "Grátis em\nconsole.groq.com"),
            ("3", "✨", "Gere a análise", "IA analisa tudo\ne gera insights"),
        ]
        for n, ico, titulo, sub in passos:
            c = tk.Frame(cards_frm, bg=COR_BRANCO,
                         highlightbackground=COR_BORDA, highlightthickness=1)
            c.pack(side="left", padx=8, ipadx=16, ipady=12)
            tk.Label(c, text=n, font=("Helvetica", 9, "bold"),
                     bg=COR_AZUL, fg=COR_BRANCO, width=2).pack(anchor="ne", padx=6, pady=6)
            tk.Label(c, text=ico, font=("Helvetica", 28), bg=COR_BRANCO).pack()
            tk.Label(c, text=titulo, font=("Helvetica", 10, "bold"),
                     bg=COR_BRANCO, fg=COR_TEXTO).pack(pady=(2, 0))
            tk.Label(c, text=sub, font=("Helvetica", 9),
                     bg=COR_BRANCO, fg=COR_MUTED, justify="center").pack(pady=(0, 8))

    # ── Importar arquivo(s) (acumula) ───────────────────────────────────────────
    def _importar(self):
        # askopenfilenames permite selecionar VÁRIOS arquivos de uma vez (Ctrl/Shift+clique)
        paths = filedialog.askopenfilenames(
            title="Selecione relatório(s) do LinkedIn (pode escolher vários)",
            filetypes=[
                ("Excel/CSV", "*.xlsx *.xlsm *.xls *.csv"),
                ("Excel", "*.xlsx *.xlsm *.xls"),
                ("CSV", "*.csv"),
                ("Todos os arquivos", "*.*"),
            ]
        )
        if not paths:
            return

        erros = []
        importados = 0

        for path in paths:
            nome = os.path.basename(path)
            ext = os.path.splitext(path)[1].lower()
            try:
                if ext == ".csv":
                    # tenta detectar separador automaticamente e cai para latin-1 se necessário
                    # (exports do LinkedIn em PT-BR às vezes usam ; e encoding latin-1)
                    try:
                        df = pd.read_csv(path, sep=None, engine="python")
                    except UnicodeDecodeError:
                        df = pd.read_csv(path, sep=None, engine="python", encoding="latin-1")
                    chave = nome.rsplit(".", 1)[0]
                    self.dados[chave] = df

                elif ext in (".xlsx", ".xlsm"):
                    xls = pd.ExcelFile(path, engine="openpyxl")
                    for sheet in xls.sheet_names:
                        chave = f"{nome.rsplit('.', 1)[0]} — {sheet}"
                        self.dados[chave] = xls.parse(sheet)

                elif ext == ".xls":
                    # arquivo .xls antigo precisa do engine xlrd (NÃO do openpyxl,
                    # que só sabe ler .xlsx — essa era a causa do "não deixa anexar excel")
                    try:
                        xls = pd.ExcelFile(path, engine="xlrd")
                    except ImportError:
                        erros.append(
                            f"{nome}: arquivo .xls antigo requer o pacote 'xlrd'.\n"
                            f"   Rode: pip install xlrd  (ou salve o arquivo como .xlsx no Excel e tente de novo)"
                        )
                        continue
                    for sheet in xls.sheet_names:
                        chave = f"{nome.rsplit('.', 1)[0]} — {sheet}"
                        self.dados[chave] = xls.parse(sheet)

                else:
                    erros.append(f"{nome}: formato não suportado ({ext}).")
                    continue

                if nome not in self.arquivos_nome:
                    self.arquivos_nome.append(nome)
                importados += 1

            except Exception as e:
                erros.append(f"{nome}: {e}")

        self._atualizar_lista_arquivos()

        total = sum(len(d) for d in self.dados.values())
        if importados:
            self.label_status.config(
                text=f"{len(self.arquivos_nome)} arquivo(s) · {total} linhas carregadas.",
                fg=COR_VERDE)
        if erros:
            self.label_status.config(text="Alguns arquivos não puderam ser lidos.", fg=COR_VERM)
            messagebox.showerror(
                "Erro ao importar",
                "Não foi possível ler o(s) seguinte(s) arquivo(s):\n\n" + "\n\n".join(erros)
            )

    # ── Limpar ─────────────────────────────────────────────────────────────────
    def _limpar(self):
        self.dados.clear()
        self.arquivos_nome.clear()
        self.texto_manual.delete("1.0", "end")
        self._atualizar_lista_arquivos()
        self.label_status.config(text="Dados limpos.", fg=COR_MUTED)
        self._tela_inicial()

    # ── Analisar ───────────────────────────────────────────────────────────────
    def _analisar(self):
        key = self.api_key.get().strip()
        if not key:
            messagebox.showwarning("API Key", "Cole sua Groq API Key no campo acima.")
            return
        manual = self.texto_manual.get("1.0", "end").strip()
        if not self.dados and not manual:
            messagebox.showwarning("Dados", "Importe ao menos um arquivo ou cole dados manualmente.")
            return

        self.btn_analisar.config(state="disabled", text="⏳  Analisando...")
        self.label_status.config(text="Enviando para a IA...", fg=COR_AZUL)
        self._mostrar_loading()

        def tarefa():
            try:
                summary = preparar_resumo(self.dados) if self.dados else ""
                if manual:
                    summary += f"\n\nDADOS ADICIONAIS / COLADOS:\n{manual}"
                resultado = analisar_com_ia(summary, key)
                self.after(0, lambda: self._mostrar_resultado(resultado))
            except Exception as e:
                msg = str(e)
                self.after(0, lambda m=msg: self._erro(m))

        threading.Thread(target=tarefa, daemon=True).start()

    # ── Loading ────────────────────────────────────────────────────────────────
    def _mostrar_loading(self):
        for w in self.area_principal.winfo_children():
            w.destroy()
        frm = tk.Frame(self.area_principal, bg=COR_CINZA)
        frm.place(relx=0.5, rely=0.45, anchor="center")
        tk.Label(frm, text="⏳", font=("Helvetica", 48), bg=COR_CINZA).pack()
        tk.Label(frm, text="Analisando seus dados...",
                 font=("Helvetica", 15, "bold"), bg=COR_CINZA, fg=COR_AZUL).pack(pady=(8, 4))
        tk.Label(frm, text="A IA está lendo todos os relatórios e gerando\ninsights personalizados para a Valtec.",
                 font=("Helvetica", 11), bg=COR_CINZA, fg=COR_MUTED, justify="center").pack()

    # ── Erro ───────────────────────────────────────────────────────────────────
    def _erro(self, msg):
        self.btn_analisar.config(state="normal", text="✨  Gerar Análise Completa")
        self.label_status.config(text="Erro na análise.", fg=COR_VERM)
        self._tela_inicial()
        m = msg.lower()
        if "api_key_invalid" in m or "invalid api key" in m:
            titulo, texto = "API Key inválida", "Acesse console.groq.com e gere uma nova chave."
        elif "quota" in m or "429" in msg:
            titulo, texto = "Cota esgotada", "Aguarde ou crie uma nova API Key em console.groq.com"
        elif "json" in m or "parse" in m:
            titulo, texto = "Erro ao processar resposta", "O modelo retornou resposta inesperada. Tente novamente."
        elif "network" in m or "connection" in m:
            titulo, texto = "Erro de conexão", "Verifique sua internet e tente novamente."
        else:
            titulo, texto = "Erro na análise", f"Detalhe técnico:\n{msg}"
        messagebox.showerror(titulo, texto)

    # ══════════════════════════════════════════════════════════════════════════
    # MOSTRAR RESULTADO
    # ══════════════════════════════════════════════════════════════════════════
    def _mostrar_resultado(self, r: dict):
        self.btn_analisar.config(state="normal", text="✨  Gerar Análise Completa")
        self.label_status.config(text="✓ Análise concluída!", fg=COR_VERDE)

        for w in self.area_principal.winfo_children():
            w.destroy()

        # Scroll container
        cs = tk.Canvas(self.area_principal, bg=COR_CINZA, highlightthickness=0)
        sb = ttk.Scrollbar(self.area_principal, orient="vertical", command=cs.yview)
        cs.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        cs.pack(side="left", fill="both", expand=True)
        cont = tk.Frame(cs, bg=COR_CINZA)
        win_id = cs.create_window((0, 0), window=cont, anchor="nw")
        cs.bind("<Configure>", lambda e: cs.itemconfig(win_id, width=e.width))
        cont.bind("<Configure>", lambda e: cs.configure(scrollregion=cs.bbox("all")))
        cs.bind_all("<MouseWheel>", lambda e: cs.yview_scroll(-1*(e.delta//120), "units"))

        # ── 1. Header com score geral ──────────────────────────────────────
        score = r.get("score_geral", 0)
        cor_score = COR_VERDE if score >= 70 else COR_AMARELO if score >= 45 else COR_VERM
        hdr_frm = tk.Frame(cont, bg=COR_AZUL_ESC, pady=16)
        hdr_frm.pack(fill="x", padx=0, pady=(0, 6))

        left_hdr = tk.Frame(hdr_frm, bg=COR_AZUL_ESC)
        left_hdr.pack(side="left", padx=20)
        tk.Label(left_hdr, text="SCORE GERAL DA PÁGINA",
                 font=("Helvetica", 9, "bold"), bg=COR_AZUL_ESC, fg="#90CAF9").pack(anchor="w")
        score_row = tk.Frame(left_hdr, bg=COR_AZUL_ESC)
        score_row.pack(anchor="w")
        tk.Label(score_row, text=str(score), font=("Helvetica", 42, "bold"),
                 bg=COR_AZUL_ESC, fg=cor_score).pack(side="left")
        tk.Label(score_row, text="/100", font=("Helvetica", 16),
                 bg=COR_AZUL_ESC, fg="#90CAF9").pack(side="left", pady=(18, 0))
        tk.Label(left_hdr, text=r.get("score_interpretacao", ""),
                 font=("Helvetica", 10), bg=COR_AZUL_ESC, fg="#DBEAFE").pack(anchor="w", pady=(2, 0))

        right_hdr = tk.Frame(hdr_frm, bg=COR_AZUL_ESC)
        right_hdr.pack(side="left", padx=30)
        tk.Label(right_hdr, text="📋 RESUMO DO PERÍODO",
                 font=("Helvetica", 9, "bold"), bg=COR_AZUL_ESC, fg="#90CAF9").pack(anchor="w")
        tk.Label(right_hdr, text=r.get("periodo", ""),
                 font=("Helvetica", 10, "bold"), bg=COR_AZUL_ESC, fg=COR_BRANCO).pack(anchor="w")
        tk.Label(right_hdr, text=r.get("resumo", ""),
                 font=("Helvetica", 10), bg=COR_AZUL_ESC, fg="#DBEAFE",
                 wraplength=500, justify="left").pack(anchor="w", pady=(4, 0))

        # ── 2. Métricas ────────────────────────────────────────────────────
        self._titulo_secao(cont, "📊 Métricas do Período")
        grid_m = tk.Frame(cont, bg=COR_CINZA)
        grid_m.pack(fill="x", padx=4, pady=(0, 4))
        cores_st = {"good": COR_VERDE, "warn": COR_AMARELO, "bad": COR_VERM}
        for i, m in enumerate(r.get("metrics", [])):
            col = i % 4
            row = i // 4
            card = tk.Frame(grid_m, bg=COR_BRANCO,
                            highlightbackground=COR_BORDA, highlightthickness=1)
            card.grid(row=row, column=col, padx=4, pady=4, sticky="nsew")
            grid_m.columnconfigure(col, weight=1)
            cor = cores_st.get(m.get("status", "warn"), COR_MUTED)
            tk.Label(card, text=m["label"], font=FONTES["small"],
                     bg=COR_BRANCO, fg=COR_MUTED).pack(pady=(10, 2), padx=12, anchor="w")
            val_row = tk.Frame(card, bg=COR_BRANCO)
            val_row.pack(padx=12, anchor="w", pady=(0, 2))
            tk.Label(val_row, text=m["value"], font=FONTES["metric"],
                     bg=COR_BRANCO, fg=cor).pack(side="left")
            var = m.get("variacao", "")
            if var:
                cor_var = COR_VERDE if "+" in var else COR_VERM if "-" in var else COR_MUTED
                tk.Label(val_row, text=f" {var}", font=("Helvetica", 9),
                         bg=COR_BRANCO, fg=cor_var).pack(side="left", pady=(8, 0))
            tk.Frame(card, bg=cor, height=3).pack(fill="x", pady=(6, 0))

        # ── 3. Score por pilar ─────────────────────────────────────────────
        self._titulo_secao(cont, "🏛 Score por Pilar")
        pilares_frm = tk.Frame(cont, bg=COR_CINZA)
        pilares_frm.pack(fill="x", padx=4, pady=(0, 4))
        for i, p in enumerate(r.get("score_por_pilar", [])):
            col = i % 5
            card = tk.Frame(pilares_frm, bg=COR_BRANCO,
                            highlightbackground=COR_BORDA, highlightthickness=1)
            card.grid(row=0, column=col, padx=4, pady=2, sticky="nsew", ipadx=8, ipady=8)
            pilares_frm.columnconfigure(col, weight=1)
            sc = p.get("score", 0)
            cor_p = COR_VERDE if sc >= 70 else COR_AMARELO if sc >= 45 else COR_VERM
            tk.Label(card, text=p["pilar"], font=("Helvetica", 9, "bold"),
                     bg=COR_BRANCO, fg=COR_TEXTO).pack(pady=(8, 2))
            tk.Label(card, text=str(sc), font=("Helvetica", 22, "bold"),
                     bg=COR_BRANCO, fg=cor_p).pack()
            # Barra de progresso visual
            bar_bg = tk.Frame(card, bg=COR_CINZA2, height=6)
            bar_bg.pack(fill="x", padx=8, pady=(2, 4))
            bar_bg.update_idletasks()
            w_bar = max(4, int(bar_bg.winfo_reqwidth() * sc / 100))
            tk.Frame(bar_bg, bg=cor_p, height=6, width=w_bar).place(x=0, y=0)
            tk.Label(card, text=p.get("descricao", ""), font=("Helvetica", 8),
                     bg=COR_BRANCO, fg=COR_MUTED, wraplength=120, justify="center").pack(padx=6, pady=(0, 8))

        # ── 4. Gráficos: dias + horários ──────────────────────────────────
        self._titulo_secao(cont, "🕐 Melhor Momento para Postar")
        frm_g1 = tk.Frame(cont, bg=COR_CINZA)
        frm_g1.pack(fill="x", padx=4, pady=(0, 4))

        fig = Figure(figsize=(12, 3.4), facecolor=COR_CINZA)

        ax1 = fig.add_subplot(1, 2, 1)
        dias_data = r.get("melhores_dias", [])
        dias = [d["dia"] for d in dias_data]
        sc_d = [d["score"] for d in dias_data]
        cores_d = [COR_VERDE if s >= 75 else COR_AZUL if s >= 45 else "#CBD5E1" for s in sc_d]
        bars = ax1.bar(dias, sc_d, color=cores_d, width=0.55, edgecolor="none")
        for bar, val in zip(bars, sc_d):
            ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                     str(val), ha="center", va="bottom", fontsize=8, color=COR_TEXTO)
        ax1.set_title("Melhores dias para postar", fontsize=10, fontweight="bold",
                      color=COR_TEXTO, pad=10)
        ax1.set_ylim(0, 120); ax1.set_facecolor(COR_BRANCO)
        ax1.tick_params(colors=COR_MUTED, labelsize=9)
        ax1.spines[:].set_visible(False); ax1.yaxis.set_visible(False)

        ax2 = fig.add_subplot(1, 2, 2)
        hor_data = r.get("melhores_horarios", [])
        horas = [h["hora"] for h in hor_data]
        sc_h  = [h["score"] for h in hor_data]
        cores_h = [COR_AZUL if s >= 75 else "#90CAF9" if s >= 45 else "#CBD5E1" for s in sc_h]
        bars2 = ax2.barh(horas, sc_h, color=cores_h, height=0.55, edgecolor="none")
        for bar, val in zip(bars2, sc_h):
            ax2.text(bar.get_width() + 1, bar.get_y() + bar.get_height()/2,
                     str(val), va="center", fontsize=8, color=COR_TEXTO)
        ax2.set_title("Melhores horários para postar", fontsize=10, fontweight="bold",
                      color=COR_TEXTO, pad=10)
        ax2.set_xlim(0, 120); ax2.set_facecolor(COR_BRANCO)
        ax2.tick_params(colors=COR_MUTED, labelsize=9)
        ax2.spines[:].set_visible(False); ax2.xaxis.set_visible(False); ax2.invert_yaxis()

        fig.tight_layout(pad=1.8)
        FigureCanvasTkAgg(fig, master=frm_g1).get_tk_widget().pack(fill="x", expand=True)
        FigureCanvasTkAgg(fig, master=frm_g1).draw()

        # ── 5. Tipos de conteúdo ──────────────────────────────────────────
        self._titulo_secao(cont, "📝 Desempenho por Tipo de Conteúdo")
        frm_g2 = tk.Frame(cont, bg=COR_CINZA)
        frm_g2.pack(fill="x", padx=4, pady=(0, 4))

        tipos_data = r.get("tipos_conteudo", [])
        tipos    = [t["tipo"]        for t in tipos_data]
        alcances = [t["alcance"]     for t in tipos_data]
        engs     = [t["engajamento"] for t in tipos_data]
        freqs    = [t.get("frequencia", 0) for t in tipos_data]

        fig2 = Figure(figsize=(12, 3.0), facecolor=COR_CINZA)
        for idx, (vals, titulo_g, cor_g) in enumerate([
            (alcances, "Alcance médio", COR_AZUL),
            (engs,     "Engajamento %", COR_VERDE),
            (freqs,    "Frequência (posts)", COR_AMARELO),
        ]):
            ax = fig2.add_subplot(1, 3, idx+1)
            ax.bar(tipos, vals, color=cor_g, width=0.5, edgecolor="none")
            ax.set_title(titulo_g, fontsize=9, fontweight="bold", color=COR_TEXTO, pad=8)
            ax.set_facecolor(COR_BRANCO)
            ax.tick_params(colors=COR_MUTED, labelsize=8)
            ax.spines[:].set_visible(False)

        fig2.tight_layout(pad=1.5)
        FigureCanvasTkAgg(fig2, master=frm_g2).get_tk_widget().pack(fill="x", expand=True)
        FigureCanvasTkADD = FigureCanvasTkAgg(fig2, master=frm_g2); FigureCanvasTkADD.draw()

        # ── 6. Audiência ──────────────────────────────────────────────────
        aud = r.get("audiencia", {})
        if aud:
            self._titulo_secao(cont, "👥 Perfil da Audiência")
            aud_frm = tk.Frame(cont, bg=COR_CINZA)
            aud_frm.pack(fill="x", padx=4, pady=(0, 4))

            for col_idx, (chave, label) in enumerate([
                ("cargos_top",        "Cargos mais comuns"),
                ("setores_top",       "Setores"),
                ("origem_seguidores", "Origem dos seguidores"),
            ]):
                items = aud.get(chave, [])
                if not items: continue
                card = tk.Frame(aud_frm, bg=COR_BRANCO,
                                highlightbackground=COR_BORDA, highlightthickness=1)
                card.grid(row=0, column=col_idx, padx=4, pady=2, sticky="nsew")
                aud_frm.columnconfigure(col_idx, weight=1)
                tk.Label(card, text=label, font=("Helvetica", 9, "bold"),
                         bg=COR_AZUL_CLR, fg=COR_AZUL_ESC).pack(fill="x", padx=0, pady=0, ipady=6)
                for it in items:
                    row_it = tk.Frame(card, bg=COR_BRANCO)
                    row_it.pack(fill="x", padx=10, pady=2)
                    nome_it = it.get("cargo") or it.get("setor") or it.get("origem", "")
                    pct = it.get("percentual", 0)
                    tk.Label(row_it, text=nome_it, font=FONTES["small"],
                             bg=COR_BRANCO, fg=COR_TEXTO).pack(side="left")
                    tk.Label(row_it, text=f"{pct}%", font=("Helvetica", 9, "bold"),
                             bg=COR_BRANCO, fg=COR_AZUL).pack(side="right")
                    # mini barra
                    bar_bg2 = tk.Frame(card, bg=COR_CINZA2, height=4)
                    bar_bg2.pack(fill="x", padx=10, pady=(0, 3))
                    bar_bg2.update_idletasks()
                    w2 = max(3, int(160 * pct / 100))
                    tk.Frame(bar_bg2, bg=COR_AZUL, height=4, width=w2).place(x=0, y=0)
                tk.Label(card, text="", bg=COR_BRANCO).pack(pady=4)

        # ── 7. Top Posts ──────────────────────────────────────────────────
        self._titulo_secao(cont, "🏆 Análise dos Melhores e Piores Posts")
        for p in r.get("top_posts", []):
            cor_borda = COR_VERDE if p["status"] == "top" else COR_VERM
            cor_bg    = COR_VERDE_CLR if p["status"] == "top" else COR_VERM_CLR
            tag = "⭐ Post Destaque" if p["status"] == "top" else "📉 Baixo Desempenho"
            card = tk.Frame(cont, bg=COR_BRANCO,
                            highlightbackground=cor_borda, highlightthickness=2)
            card.pack(fill="x", padx=4, pady=4)
            # Header do card
            card_hdr = tk.Frame(card, bg=cor_bg)
            card_hdr.pack(fill="x")
            tk.Label(card_hdr, text=p.get("titulo", ""), font=("Helvetica", 10, "bold"),
                     bg=cor_bg, fg=COR_TEXTO, wraplength=680, justify="left").pack(
                     side="left", padx=12, pady=8)
            tk.Label(card_hdr, text=tag, font=("Helvetica", 9, "bold"),
                     bg=cor_borda, fg=COR_BRANCO, padx=8, pady=4).pack(side="right", padx=8, pady=6)
            # Métricas
            meta_row = tk.Frame(card, bg=COR_BRANCO)
            meta_row.pack(fill="x", padx=12, pady=6)
            tipo_post = p.get("tipo", "")
            if tipo_post:
                tk.Label(meta_row, text=f"📌 {tipo_post}", font=FONTES["small"],
                         bg=COR_AZUL_CLR, fg=COR_AZUL, padx=6, pady=2).pack(side="left", padx=(0, 10))
            tk.Label(meta_row, text=f"👁 {int(p.get('impressoes', 0)):,} impressões",
                     font=FONTES["small"], bg=COR_BRANCO, fg=COR_MUTED).pack(side="left", padx=(0, 12))
            tk.Label(meta_row, text=f"❤️ {p.get('engajamento', 0):.1f}% engajamento",
                     font=FONTES["small"], bg=COR_BRANCO, fg=COR_MUTED).pack(side="left", padx=(0, 12))
            cliques = p.get("cliques", 0)
            if cliques:
                tk.Label(meta_row, text=f"🔗 {cliques} cliques",
                         font=FONTES["small"], bg=COR_BRANCO, fg=COR_MUTED).pack(side="left")
            # Padrão + motivo
            if p.get("padrao"):
                tk.Label(card, text=f"📐 Padrão: {p['padrao']}",
                         font=("Helvetica", 9, "bold"), bg=COR_BRANCO, fg=COR_TEXTO,
                         wraplength=780, justify="left").pack(padx=12, pady=(0, 2), anchor="w")
            tk.Label(card, text=p.get("motivo", ""),
                     font=FONTES["small"], bg=COR_BRANCO, fg=COR_MUTED,
                     wraplength=780, justify="left").pack(padx=12, pady=(0, 10), anchor="w")

        # ── 8. Padrões de Sucesso ─────────────────────────────────────────
        padroes = r.get("padroes_sucesso", [])
        if padroes:
            self._titulo_secao(cont, "🔬 Padrões dos Posts de Sucesso")
            for p in padroes:
                imp = p.get("impacto", "medio")
                cor_imp = COR_VERDE if imp == "alto" else COR_AMARELO if imp == "medio" else COR_MUTED
                card = tk.Frame(cont, bg=COR_BRANCO,
                                highlightbackground=cor_imp, highlightthickness=1)
                card.pack(fill="x", padx=4, pady=3)
                row_p = tk.Frame(card, bg=COR_BRANCO)
                row_p.pack(fill="x", padx=12, pady=8)
                tk.Label(row_p, text=f"Impacto {imp.upper()}",
                         font=("Helvetica", 8, "bold"), bg=cor_imp, fg=COR_BRANCO,
                         padx=6, pady=2).pack(side="right")
                tk.Label(row_p, text=f"📐 {p.get('padrao', '')}",
                         font=("Helvetica", 10, "bold"), bg=COR_BRANCO, fg=COR_TEXTO).pack(side="left", anchor="w")
                ex = p.get("exemplo", "")
                if ex:
                    tk.Label(card, text=f"Exemplo: {ex}",
                             font=FONTES["small"], bg=COR_BRANCO, fg=COR_MUTED,
                             wraplength=780, justify="left").pack(padx=12, pady=(0, 8), anchor="w")

        # ── 9. Insights ───────────────────────────────────────────────────
        self._titulo_secao(cont, "🔍 O que está Funcionando e o que Precisa Melhorar")
        cores_ins = {
            "erro": (COR_VERM_CLR, COR_VERM, "🔴"),
            "ok":   (COR_VERDE_CLR, COR_VERDE, "🟢"),
            "dica": (COR_AZUL_CLR,  COR_AZUL,  "💡"),
        }
        for ins in r.get("insights", []):
            bg_i, brd_i, ico = cores_ins.get(ins.get("tipo", "dica"), (COR_CINZA, COR_BORDA, "•"))
            imp_i = ins.get("impacto", "")
            card = tk.Frame(cont, bg=bg_i,
                            highlightbackground=brd_i, highlightthickness=1)
            card.pack(fill="x", padx=4, pady=3)
            top_row = tk.Frame(card, bg=bg_i)
            top_row.pack(fill="x", padx=12, pady=(8, 2))
            tk.Label(top_row, text=f"{ico}  {ins.get('titulo', '')}",
                     font=("Helvetica", 10, "bold"), bg=bg_i, fg=COR_TEXTO).pack(side="left")
            if imp_i:
                cor_imp2 = COR_VERM if imp_i == "alto" else COR_AMARELO
                tk.Label(top_row, text=f"Impacto {imp_i}",
                         font=("Helvetica", 8), bg=cor_imp2, fg=COR_BRANCO,
                         padx=5, pady=1).pack(side="right")
            tk.Label(card, text=ins.get("texto", ""), font=FONTES["small"],
                     bg=bg_i, fg=COR_MUTED, wraplength=780, justify="left").pack(
                     padx=12, pady=(0, 8), anchor="w")

        # ── 10. Calendário Editorial ──────────────────────────────────────
        cal = r.get("calendario_sugerido", [])
        if cal:
            self._titulo_secao(cont, "📅 Calendário Editorial Sugerido")
            for semana in cal:
                sem_frm = tk.Frame(cont, bg=COR_BRANCO,
                                   highlightbackground=COR_BORDA, highlightthickness=1)
                sem_frm.pack(fill="x", padx=4, pady=4)
                tk.Label(sem_frm, text=semana.get("semana", ""),
                         font=("Helvetica", 10, "bold"),
                         bg=COR_AZUL_ESC, fg=COR_BRANCO).pack(fill="x", ipady=6, padx=0)
                for post_cal in semana.get("posts", []):
                    row_c = tk.Frame(sem_frm, bg=COR_BRANCO)
                    row_c.pack(fill="x", padx=12, pady=5)
                    tag_dia = tk.Label(row_c,
                                       text=f"{post_cal.get('dia','')}  {post_cal.get('horario','')}",
                                       font=("Helvetica", 9, "bold"),
                                       bg=COR_AZUL, fg=COR_BRANCO, padx=8, pady=3)
                    tag_dia.pack(side="left")
                    tag_tipo = tk.Label(row_c, text=post_cal.get("tipo", ""),
                                        font=("Helvetica", 9),
                                        bg=COR_ROXO_CLR, fg=COR_ROXO, padx=6, pady=3)
                    tag_tipo.pack(side="left", padx=(6, 0))
                    tk.Label(row_c, text=post_cal.get("tema", ""),
                             font=("Helvetica", 10, "bold"),
                             bg=COR_BRANCO, fg=COR_TEXTO).pack(side="left", padx=10)
                    just = post_cal.get("justificativa", "")
                    if just:
                        tk.Label(row_c, text=f"— {just}", font=FONTES["small"],
                                 bg=COR_BRANCO, fg=COR_MUTED).pack(side="left")
                tk.Label(sem_frm, text="", bg=COR_BRANCO).pack(pady=2)

        # ── 11. Metas 30 dias ─────────────────────────────────────────────
        metas = r.get("metas_30dias", [])
        if metas:
            self._titulo_secao(cont, "🎯 Metas para os Próximos 30 Dias")
            metas_frm = tk.Frame(cont, bg=COR_CINZA)
            metas_frm.pack(fill="x", padx=4, pady=(0, 4))
            for i, meta in enumerate(metas):
                card = tk.Frame(metas_frm, bg=COR_BRANCO,
                                highlightbackground=COR_BORDA, highlightthickness=1)
                card.grid(row=0, column=i, padx=4, pady=2, sticky="nsew")
                metas_frm.columnconfigure(i, weight=1)
                tk.Label(card, text=meta.get("meta", ""), font=("Helvetica", 9, "bold"),
                         bg=COR_AZUL_CLR, fg=COR_AZUL_ESC).pack(fill="x", ipady=5)
                inner = tk.Frame(card, bg=COR_BRANCO)
                inner.pack(padx=12, pady=8, fill="x")
                tk.Label(inner, text="Atual", font=FONTES["small"], bg=COR_BRANCO, fg=COR_MUTED).pack(anchor="w")
                tk.Label(inner, text=meta.get("atual", "N/D"), font=("Helvetica", 14, "bold"),
                         bg=COR_BRANCO, fg=COR_TEXTO).pack(anchor="w")
                tk.Label(inner, text="▶ Meta", font=FONTES["small"], bg=COR_BRANCO, fg=COR_MUTED).pack(anchor="w", pady=(6,0))
                tk.Label(inner, text=meta.get("alvo", "N/D"), font=("Helvetica", 14, "bold"),
                         bg=COR_BRANCO, fg=COR_VERDE).pack(anchor="w")
                tk.Label(inner, text=meta.get("acao", ""), font=FONTES["small"],
                         bg=COR_BRANCO, fg=COR_MUTED, wraplength=180, justify="left").pack(anchor="w", pady=(6,0))

        # ── 12. Recomendações ─────────────────────────────────────────────
        self._titulo_secao(cont, "🚀 Recomendações Estratégicas para a Valtec")
        cores_prio = {"alta": COR_VERM, "media": COR_AMARELO, "baixa": COR_VERDE}
        for i, rec in enumerate(r.get("recomendacoes", []), 1):
            if isinstance(rec, str):
                # fallback se vier como string simples
                card = tk.Frame(cont, bg=COR_BRANCO,
                                highlightbackground=COR_BORDA, highlightthickness=1)
                card.pack(fill="x", padx=4, pady=3)
                tk.Label(card, text=str(i), font=("Helvetica", 12, "bold"),
                         bg=COR_AZUL, fg=COR_BRANCO, width=3, pady=10).pack(side="left", fill="y")
                tk.Label(card, text=rec, font=FONTES["body"],
                         bg=COR_BRANCO, fg=COR_TEXTO, wraplength=730, justify="left").pack(
                         side="left", padx=12, pady=10)
            else:
                prio = rec.get("prioridade", "media")
                cor_p2 = cores_prio.get(prio, COR_MUTED)
                prazo  = rec.get("prazo", "")
                card = tk.Frame(cont, bg=COR_BRANCO,
                                highlightbackground=cor_p2, highlightthickness=2)
                card.pack(fill="x", padx=4, pady=4)
                top_rec = tk.Frame(card, bg=COR_BRANCO)
                top_rec.pack(fill="x", padx=12, pady=(10, 4))
                tk.Label(top_rec, text=f"{i}. {rec.get('titulo', '')}",
                         font=("Helvetica", 11, "bold"),
                         bg=COR_BRANCO, fg=COR_TEXTO).pack(side="left")
                tag_frm = tk.Frame(top_rec, bg=COR_BRANCO)
                tag_frm.pack(side="right")
                tk.Label(tag_frm, text=f"Prioridade {prio.upper()}",
                         font=("Helvetica", 8, "bold"),
                         bg=cor_p2, fg=COR_BRANCO, padx=6, pady=2).pack(side="left", padx=2)
                if prazo:
                    tk.Label(tag_frm, text=prazo,
                             font=("Helvetica", 8),
                             bg=COR_CINZA2, fg=COR_MUTED, padx=6, pady=2).pack(side="left", padx=2)
                tk.Label(card, text=rec.get("descricao", ""),
                         font=FONTES["body"], bg=COR_BRANCO, fg=COR_MUTED,
                         wraplength=780, justify="left").pack(padx=12, pady=(0, 12), anchor="w")

        # Rodapé
        rodape = tk.Frame(cont, bg=COR_AZUL_ESC, pady=10)
        rodape.pack(fill="x", pady=(20, 0))
        tk.Label(rodape, text="Valtec LinkedIn Analyzer  •  Análise gerada por IA  •  Groq + Llama 3.3",
                 font=("Helvetica", 9), bg=COR_AZUL_ESC, fg="#90CAF9").pack()

        tk.Label(cont, text="", bg=COR_CINZA).pack(pady=10)


# ── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = ValtecApp()
    app.mainloop()