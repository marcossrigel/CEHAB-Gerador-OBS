from openai import OpenAI
import os
import re
import threading
import tkinter as tk
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional, Tuple

client = OpenAI()

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

@dataclass
class DocumentoSEI:
    caminho: Path
    nome_arquivo: str
    texto: str
    codigo_documento: Optional[str] = None
    tipo_documento: str = "DESCONHECIDO"
    data_assinatura: Optional[str] = None  # dd/mm/aaaa
    resumo: str = ""


@dataclass
class AnaliseProcesso:
    remanejamento_orcamentario: bool = False
    documentos: List[DocumentoSEI] = field(default_factory=list)
    data_solicitacao_gop: Optional[str] = None
    codigo_oficio_gop: Optional[str] = None
    possui_reiteracao: bool = False
    data_reiteracao: Optional[str] = None
    codigo_reiteracao: Optional[str] = None
    em_analise: bool = False
    orgao_atual: Optional[str] = None
    sem_manifestacao: bool = False
    destaque_realizado: bool = False
    programacao_financeira: bool = False
    desdobramento_fonte: bool = False
    autorizacao_execucao: bool = False
    dotacao_orcamentaria: bool = False
    docs_relevantes: List[str] = field(default_factory=list)
    observacoes_tecnicas: List[str] = field(default_factory=list)


def limpar_texto(texto: str) -> str:
    texto = texto.replace("\x00", " ")
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    return texto.strip()


def extrair_texto_pdf(caminho_pdf: Path) -> str:
    if PdfReader is None:
        raise RuntimeError(
            "Biblioteca 'pypdf' não encontrada. Instale com: pip install pypdf"
        )

    partes: List[str] = []
    reader = PdfReader(str(caminho_pdf))
    for pagina in reader.pages:
        try:
            partes.append(pagina.extract_text() or "")
        except Exception:
            partes.append("")
    return limpar_texto("\n".join(partes))


def limpar_linha_destinatario(linha: str) -> str:
    linha = re.sub(r"\s+", " ", linha).strip(" -:\t")
    return linha.strip()


def linha_e_tratamento(linha: str) -> bool:
    l = linha.upper().strip()
    tratamentos = {
        "À EXMA SRA.", "A EXMA SRA.", "À EXMA. SRA.", "A EXMA. SRA.",
        "EXMA. SRA.", "EXMA SRA.", "ILMA. SRA.", "ILMA SRA.",
        "À EXMO SR.", "A EXMO SR.", "EXMO. SR.", "EXMO SR.",
        "ILMO. SR.", "ILMO SR.", "AO EXCELENTÍSSIMO SENHOR",
        "AO EXCELENTISSIMO SENHOR", "À SENHORA", "A SENHORA",
        "AO SENHOR"
    }
    return l in tratamentos


def linha_parece_assinatura_rodape(linha: str) -> bool:
    l = linha.upper().strip()
    chaves = [
        "ATENCIOSAMENTE",
        "DOCUMENTO ASSINADO ELETRONICAMENTE",
        "A AUTENTICIDADE DESTE DOCUMENTO",
        "COMPANHIA ESTADUAL DE HABITAÇÃO E OBRAS",
        "RUA ODORICO MENDES",
        "GOVPE -",
    ]
    return any(chave in l for chave in chaves)


def extrair_destinatario_oficio(texto: str) -> Optional[str]:
    linhas = [limpar_linha_destinatario(l) for l in texto.split("\n") if l.strip()]

    inicio = None
    for i, linha in enumerate(linhas):
        linha_upper = linha.upper()
        if (
            re.match(r"^(À|AO|AOS|ÀS)\b", linha_upper)
            or linha_e_tratamento(linha)
        ):
            inicio = i
            break

    if inicio is None:
        return None

    bloco = []
    for linha in linhas[inicio:inicio + 10]:
        linha_upper = linha.upper()

        if linha_upper.startswith("ASSUNTO:"):
            break

        if linha_e_tratamento(linha):
            continue

        if re.match(r"^OF[IÍ]CIO\s+N[ºO]?", linha_upper):
            continue

        if linha_parece_assinatura_rodape(linha):
            break

        bloco.append(linha)

    if not bloco:
        return None

    # remove linhas muito curtas de chamamento tipo "Prezado,"
    bloco_filtrado = []
    for linha in bloco:
        linha_upper = linha.upper()
        if linha_upper in {"PREZADO,", "PREZADA,", "PREZADOS,", "PREZADAS,"}:
            continue
        bloco_filtrado.append(linha)

    if not bloco_filtrado:
        return None

    # regra prática:
    # pega até 2 linhas antes do "Assunto":
    # 1ª = nome
    # 2ª = cargo/órgão
    return " ".join(bloco_filtrado[:2]).strip()


def extrair_nome_e_cargo_destinatario(texto: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    linhas = [limpar_linha_destinatario(l) for l in texto.split("\n") if l.strip()]

    inicio = None
    for i, linha in enumerate(linhas):
        if (
            re.match(r"^(À|AO|AOS|ÀS)\b", linha.upper())
            or linha_e_tratamento(linha)
        ):
            inicio = i
            break

    if inicio is None:
        return None, None, None

    coletadas = []
    for linha in linhas[inicio:inicio + 10]:
        linha_upper = linha.upper()

        if linha_upper.startswith("ASSUNTO:"):
            break

        if linha_e_tratamento(linha):
            continue

        if re.match(r"^OF[IÍ]CIO\s+N[ºO]?", linha_upper):
            continue

        if linha_parece_assinatura_rodape(linha):
            break

        if linha_upper in {"PREZADO,", "PREZADA,", "PREZADOS,", "PREZADAS,"}:
            continue

        coletadas.append(linha)

    nome = coletadas[0] if len(coletadas) >= 1 else None
    cargo = coletadas[1] if len(coletadas) >= 2 else None
    completo = " ".join([x for x in [nome, cargo] if x]).strip() or None

    return nome, cargo, completo

def extrair_codigo_documento(nome_arquivo: str, texto: str) -> Optional[str]:
    m = re.search(r"SEI_(\d{6,9})_", nome_arquivo)
    if m:
        return m.group(1)

    m = re.search(r"c[oó]digo verificador\s+(\d{6,9})", texto, flags=re.I)
    if m:
        return m.group(1)

    m = re.search(r"\((\d{6,9})\)", texto)
    if m:
        return m.group(1)

    return None


def extrair_data_assinatura(texto: str) -> Optional[str]:
    padroes = [
        r"em\s+(\d{2}/\d{2}/\d{4}),\s+às",
        r"Recife,\s+(\d{2}\s+de\s+[A-Za-zçãéíóúâêô]+\s+de\s+\d{4})",
    ]

    for padrao in padroes:
        m = re.search(padrao, texto, flags=re.I)
        if m:
            valor = m.group(1)
            if re.match(r"\d{2}/\d{2}/\d{4}", valor):
                return valor
            return normalizar_data_extenso(valor)
    return None


def normalizar_data_extenso(valor: str) -> Optional[str]:
    meses = {
        "janeiro": "01",
        "fevereiro": "02",
        "março": "03",
        "marco": "03",
        "abril": "04",
        "maio": "05",
        "junho": "06",
        "julho": "07",
        "agosto": "08",
        "setembro": "09",
        "outubro": "10",
        "novembro": "11",
        "dezembro": "12",
    }
    m = re.search(r"(\d{1,2})\s+de\s+([A-Za-zçãéíóúâêô]+)\s+de\s+(\d{4})", valor, flags=re.I)
    if not m:
        return None
    dia, mes_ext, ano = m.groups()
    mes = meses.get(mes_ext.lower())
    if not mes:
        return None
    return f"{int(dia):02d}/{mes}/{ano}"


def detectar_tipo_documento(nome_arquivo: str, texto: str) -> str:
    nome_upper = nome_arquivo.upper()
    texto_upper = texto.upper()

    if "DESPACHO" in nome_upper:
        return "DESPACHO"
    if "OFICIO" in nome_upper or "OFÍCIO" in nome_upper:
        return "OFICIO"
    if re.search(r"(^|[^A-Z])CI([^A-Z]|$)", nome_upper) or "COMUNICAÇÃO INTERNA" in nome_upper or "COMUNICACAO INTERNA" in nome_upper:
        return "CI"
    if "AUTORIZACAO" in nome_upper or "AUTORIZAÇÃO" in nome_upper:
        return "AUTORIZACAO"
    if "SOF" in nome_upper:
        return "SOF"

    if re.search(r"\bDESPACHO\b", texto_upper):
        return "DESPACHO"
    if re.search(r"\bOF[IÍ]CIO\b", texto_upper):
        return "OFICIO"
    if re.search(r"\bCI\b|COMUNICAÇÃO INTERNA|COMUNICACAO INTERNA", texto_upper):
        return "CI"
    if "AUTORIZAÇÃO" in texto_upper or "AUTORIZACAO" in texto_upper:
        return "AUTORIZACAO"
    if "SOLICITAÇÃO ORÇAMENTÁRIA" in texto_upper or "SOLICITACAO ORCAMENTARIA" in texto_upper:
        return "SOF"

    return "OUTRO"


def extrair_destinatario(texto: str) -> Optional[str]:
    m = re.search(r"Destinat[aá]rio:\s*(.+)", texto, flags=re.I)
    if m:
        return m.group(1).strip()
    return None


def texto_contem(texto: str, termos: List[str]) -> bool:
    t = texto.upper()
    return any(term.upper() in t for term in termos)


def documento_e_oficio_gop(doc: DocumentoSEI) -> bool:
    return (
        doc.tipo_documento == "OFICIO"
        and "CEHAB/GOP" in doc.texto.upper()
    )


def documento_e_reiteracao(doc: DocumentoSEI) -> bool:
    t = doc.texto.upper()
    return documento_e_oficio_gop(doc) and (
        "REITERANDO O OFÍCIO" in t
        or "REITERAMOS" in t
        or "REITERA" in t
    )


def analisar_documentos(documentos: List[DocumentoSEI]) -> AnaliseProcesso:
    analise = AnaliseProcesso(documentos=documentos)

    # Ordenação por data, quando existir
    documentos_ordenados = sorted(
        documentos,
        key=lambda d: d.data_assinatura or "99/99/9999"
    )

    for doc in documentos_ordenados:
        texto_upper = doc.texto.upper()
        codigo = doc.codigo_documento

        # 1) Solicitação principal da GOP
        if documento_e_oficio_gop(doc):
            if analise.data_solicitacao_gop is None:
                analise.data_solicitacao_gop = doc.data_assinatura
                analise.codigo_oficio_gop = codigo
                analise.docs_relevantes.append(f"Ofício {codigo}" if codigo else "Ofício")
            if "DECLARAÇÃO DE DOTAÇÃO ORÇAMENTÁRIA" in texto_upper or "DECLARACAO DE DOTACAO ORCAMENTARIA" in texto_upper:
                analise.dotacao_orcamentaria = True

        # 2) Reiteração
        if documento_e_reiteracao(doc):
            analise.possui_reiteracao = True
            analise.data_reiteracao = doc.data_assinatura
            analise.codigo_reiteracao = codigo
            item = f"Ofício {codigo}" if codigo else "Ofício de reiteração"
            if item not in analise.docs_relevantes:
                analise.docs_relevantes.append(item)

        # 3) Tramitação / análise
        destinatario = extrair_destinatario(doc.texto)
        if destinatario:
            dest_upper = destinatario.upper()
            if "GOP" not in dest_upper and any(org in dest_upper for org in ["SEDUH", "SECTI", "SESP", "SEAP", "SEE", "ADAGRO", "SUPOF", "SEPLAG", "DPEC", "DGAF"]):
                analise.em_analise = True
                analise.orgao_atual = destinatario

        if texto_contem(doc.texto, [
            "SEGUE PARA ANÁLISE",
            "PARA ANÁLISE",
            "ENCAMINHO O OFÍCIO",
            "PARA PROVIDÊNCIAS CABÍVEIS",
            "EM TRAMITAÇÃO",
        ]):
            if analise.orgao_atual is None:
                analise.em_analise = True

        # 4) Programação financeira
        if texto_contem(doc.texto, [
            "PROGRAMAÇÃO FINANCEIRA",
            "DISPONIBILIZAÇÃO DE PROGRAMAÇÃO FINANCEIRA",
            "DISPONIBILIZACAO DE PROGRAMACAO FINANCEIRA",
            "PF ",
        ]):
            analise.programacao_financeira = True
            item = f"Despacho {codigo}" if codigo else doc.tipo_documento.title()
            if item not in analise.docs_relevantes:
                analise.docs_relevantes.append(item)

        # 5) Desdobramento de fonte
        if texto_contem(doc.texto, [
            "DESDOBRAMENTO DE FONTE",
            "DESDOBRAMENTO DA FONTE",
            "PROCEDIDO AO DESDOBRAMENTO DA FONTE",
            "FONTE DETALHADA",
        ]):
            analise.desdobramento_fonte = True
            item = f"Despacho {codigo}" if codigo else doc.tipo_documento.title()
            if item not in analise.docs_relevantes:
                analise.docs_relevantes.append(item)
        
        def rotulo_doc(doc: DocumentoSEI) -> str:
            prefixos = {
                "OFICIO": "Ofício",
                "DESPACHO": "Despacho",
                "CI": "CI",
                "AUTORIZACAO": "Autorização",
                "SOF": "SOF",
            }
            prefixo = prefixos.get(doc.tipo_documento, "Documento")
            return f"{prefixo} {doc.codigo_documento}" if doc.codigo_documento else prefixo

        # 6) Autorização / TED
        if doc.tipo_documento == "AUTORIZACAO":
            analise.autorizacao_execucao = True
            item = rotulo_doc(doc)
            if item not in analise.docs_relevantes:
                analise.docs_relevantes.append(item)

        # 7) Destaque realizado / atendido
        if texto_contem(doc.texto, [
            "FORAM REALIZADOS OS DESTAQUES",
            "DESTAQUES ORÇAMENTÁRIOS",
            "DESTAQUES ORCAMENTARIOS",
            "PLEITO FOI ATENDIDO",
            "MESMO FOI ATENDIDO",
            "Destaque orçamentário já realizado",
        ]):
            analise.destaque_realizado = True
            item = f"Despacho {codigo}" if codigo else doc.tipo_documento.title()
            if item not in analise.docs_relevantes:
                analise.docs_relevantes.append(item)
        
        # 5.1) Remanejamento orçamentário
        if texto_contem(doc.texto, [
            "REMANEJAMENTO ORÇAMENTÁRIO",
            "REMANEJAMENTO ORCAMENTARIO",
            "AUTORIZAR O REMANEJAMENTO",
            "REMANEJAMENTO 595",
        ]):
            analise.remanejamento_orcamentario = True
            item = f"{doc.tipo_documento.title()} {codigo}" if codigo else doc.tipo_documento.title()
            if item not in analise.docs_relevantes:
                analise.docs_relevantes.append(item)

    # Regra derivada de situação
    if not (analise.destaque_realizado or analise.programacao_financeira or analise.desdobramento_fonte or analise.autorizacao_execucao):
        if not analise.em_analise and analise.data_solicitacao_gop:
            analise.sem_manifestacao = True

    # Deduplicação dos docs
    vistos = set()
    docs_final = []
    for item in analise.docs_relevantes:
        if item not in vistos:
            docs_final.append(item)
            vistos.add(item)
    analise.docs_relevantes = docs_final

    return analise


def extrair_numero_oficio(texto: str) -> str:
    m = re.search(r"Of[ií]cio\s*N[ºo]?\s*(\d+/\d{4})", texto, re.I)
    if m:
        return m.group(1)
    return "[número não identificado]"


def extrair_destinatario_oficio(texto: str) -> Optional[str]:
    linhas = [linha.strip() for linha in texto.split("\n") if linha.strip()]

    for i, linha in enumerate(linhas):
        if re.search(r"^(Ao|À)\s+", linha, re.I) or re.search(r"^(Ilma\.?|Ilmo\.?|Exma\.?|Exmo\.?)", linha, re.I):
            bloco = linhas[i:i+8]

            for item in bloco:
                item_upper = item.upper()

                if item_upper.startswith("ASSUNTO:"):
                    break

                if item_upper.startswith("OFÍCIO Nº") or item_upper.startswith("OFICIO Nº"):
                    continue

                if item_upper in {
                    "ILMA. SRA.", "ILMO. SR.", "À EXMA SRA.", "AO EXCELENTÍSSIMO SENHOR",
                    "AO EXCELENTISSIMO SENHOR", "À EXMA.", "EXMA. SRA.", "EXMO. SR."
                }:
                    continue

                if re.fullmatch(r"[A-ZÁÉÍÓÚÃÕÇ\s]+", item) and len(item.split()) >= 2:
                    continue

                return item

    return None


def normalizar_orgao_destino(destinatario: Optional[str]) -> str:
    if not destinatario:
        return "órgão não identificado"

    d = destinatario.upper()

    mapa = [
        ("SECRETARIA DE DESENVOLVIMENTO URBANO E HABITAÇÃO", "SEDUH"),
        ("SECRETARIA DE DESENVOLVIMENTO URBANO E HABITACAO", "SEDUH"),
        ("SECRETÁRIA DE SAÚDE DO ESTADO DE PERNAMBUCO", "SES-PE"),
        ("SECRETARIA DE SAÚDE DO ESTADO DE PERNAMBUCO", "SES-PE"),
        ("SECRETARIA DE SAUDE DO ESTADO DE PERNAMBUCO", "SES-PE"),
        ("SEGI/SDS", "SEGI/SDS"),
        ("SECRETÁRIO EXECUTIVO DE GESTÃO INTEGRADA - SEGI/SDS", "SEGI/SDS"),
        ("SECRETARIO EXECUTIVO DE GESTAO INTEGRADA - SEGI/SDS", "SEGI/SDS"),
        ("SECTI", "SECTI"),
        ("SEPLAG", "SEPLAG"),
        ("SEFAZ", "SEFAZ"),
        ("SDS", "SDS"),
        ("SEE", "SEE"),
        ("SES", "SES"),
        ("SECULT", "SECULT"),
        ("SEAP", "SEAP"),
        ("SESP", "SESP"),
        ("UPE", "UPE"),
        ("SETUR", "SETUR"),
        ("SECMULHER", "SECMULHER"),
        ("ADAGRO", "ADAGRO"),
        ("HEMOPE", "HEMOPE"),
    ]

    for chave, sigla in mapa:
        if chave in d:
            return sigla

    m = re.search(r"\b([A-Z]{2,}(?:/[A-Z]{2,})+)\b", destinatario)
    if m:
        return m.group(1)

    return destinatario


def identificar_retorno_para_gop(documentos: List[DocumentoSEI]) -> Optional[DocumentoSEI]:
    for doc in documentos:
        texto_upper = doc.texto.upper()

        if "DESTINATÁRIO: CEHAB GOP" in texto_upper or "DESTINATÁRIO: CEHAB/GOP" in texto_upper:
            return doc

        if "CEHAB GOP" in texto_upper or "CEHAB/GOP" in texto_upper:
            if any(chave in texto_upper for chave in [
                "ENCAMINHO",
                "ENCAMINHAMOS",
                "SEGUE",
                "EM RESPOSTA",
                "PARA PROVIDÊNCIAS",
                "PARA ANÁLISE",
            ]):
                return doc

    return None


def extrair_ano(texto: str) -> str:
    m = re.search(r"exerc[ií]cio\s+de\s+(\d{4})", texto, re.I)
    if m:
        return m.group(1)
    return "2026"

def identificar_oficio_principal(documentos: List[DocumentoSEI]) -> Optional[DocumentoSEI]:
    oficios_gop = [
        doc for doc in documentos
        if doc.tipo_documento == "OFICIO" and "CEHAB/GOP" in doc.texto.upper()
    ]

    if not oficios_gop:
        return None

    nao_reiterados = [doc for doc in oficios_gop if not documento_e_reiteracao(doc)]
    if nao_reiterados:
        return sorted(
            nao_reiterados,
            key=lambda d: d.data_assinatura or "99/99/9999"
        )[0]

    return sorted(
        oficios_gop,
        key=lambda d: d.data_assinatura or "99/99/9999"
    )[0]

def gerar_obs(analise: AnaliseProcesso, documentos: List[DocumentoSEI]) -> str:
    oficio = identificar_oficio_principal(documentos)
    reiteracao = identificar_reiteracao(documentos)

    if not oficio:
        return "Não foi identificado Ofício da GOP."

    numero_oficio = extrair_numero_oficio(oficio.texto)
    codigo = oficio.codigo_documento or "-"
    data = oficio.data_assinatura or "[sem data]"
    nome_dest, cargo_dest, destinatario_completo = extrair_nome_e_cargo_destinatario(oficio.texto)
    ano = extrair_ano(oficio.texto)

    destino_obs = destinatario_completo or "destinatário não identificado"

    obs = (
        f"Ofício Nº {numero_oficio} (Doc. SEI Nº {codigo}) "
        f"encaminhado a {destino_obs} em {data} "
        f"solicitando destaque orçamentário referente ao exercício de {ano}."
    )

    if reiteracao and reiteracao != oficio:
        numero_r = extrair_numero_oficio(reiteracao.texto) if reiteracao.tipo_documento == "OFICIO" else "-"
        data_r = reiteracao.data_assinatura or "[sem data]"
        codigo_r = reiteracao.codigo_documento or "-"

        if reiteracao.tipo_documento == "OFICIO":
            obs += (
                f" Reiterado em {data_r} através do Ofício Nº {numero_r} "
                f"(Doc. SEI Nº {codigo_r})."
            )
        else:
            obs += (
                f" Reiterado em {data_r} através do {reiteracao.tipo_documento.title()} "
                f"(Doc. SEI Nº {codigo_r})."
            )

    return obs

def identificar_reiteracao(documentos: List[DocumentoSEI]) -> Optional[DocumentoSEI]:
    oficios_reiterados = [
        doc for doc in documentos
        if doc.tipo_documento == "OFICIO"
        and "CEHAB/GOP" in doc.texto.upper()
        and documento_e_reiteracao(doc)
    ]

    if oficios_reiterados:
        return sorted(
            oficios_reiterados,
            key=lambda d: d.data_assinatura or "99/99/9999"
        )[0]

    despachos_reiterados = [
        doc for doc in documentos
        if doc.tipo_documento == "DESPACHO" and "REITER" in doc.texto.upper()
    ]

    if despachos_reiterados:
        return sorted(
            despachos_reiterados,
            key=lambda d: d.data_assinatura or "99/99/9999"
        )[0]

    return None

def carregar_documentos(caminhos: List[str]) -> List[DocumentoSEI]:
    documentos: List[DocumentoSEI] = []

    for caminho_str in caminhos:
        caminho = Path(caminho_str)
        texto = extrair_texto_pdf(caminho)
        codigo = extrair_codigo_documento(caminho.name, texto)
        tipo = detectar_tipo_documento(caminho.name, texto)
        data = extrair_data_assinatura(texto)

        documentos.append(
            DocumentoSEI(
                caminho=caminho,
                nome_arquivo=caminho.name,
                texto=texto,
                codigo_documento=codigo,
                tipo_documento=tipo,
                data_assinatura=data,
            )
        )

    return documentos


class AppSEIObs:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Leitor SEI - Geração de OBS GOP")
        self.root.geometry("1080x760")
        self.root.minsize(980, 680)

        self.caminhos_arquivos: List[str] = []
        self._montar_interface()

    def _montar_interface(self) -> None:
        estilo = ttk.Style()
        try:
            estilo.theme_use("clam")
        except Exception:
            pass

        frame_topo = ttk.Frame(self.root, padding=12)
        frame_topo.pack(fill="x")

        titulo = ttk.Label(
            frame_topo,
            text="Gerador de OBS - GOP a partir de PDFs do SEI",
            font=("Segoe UI", 16, "bold"),
        )
        titulo.pack(anchor="w")

        subtitulo = ttk.Label(
            frame_topo,
            text="Selecione os PDFs do mesmo processo, processe e copie a OBS padronizada.",
            font=("Segoe UI", 10),
        )
        subtitulo.pack(anchor="w", pady=(4, 0))

        frame_botoes = ttk.Frame(self.root, padding=(12, 0, 12, 8))
        frame_botoes.pack(fill="x")

        ttk.Button(frame_botoes, text="Selecionar PDFs", command=self.selecionar_arquivos).pack(side="left")
        ttk.Button(frame_botoes, text="Limpar", command=self.limpar).pack(side="left", padx=8)
        ttk.Button(frame_botoes, text="Processar", command=self.processar).pack(side="left")
        ttk.Button(frame_botoes, text="Copiar OBS", command=self.copiar_obs).pack(side="left", padx=8)

        self.label_status = ttk.Label(frame_botoes, text="Nenhum arquivo selecionado.")
        self.label_status.pack(side="right")

        corpo = ttk.Panedwindow(self.root, orient="horizontal")
        corpo.pack(fill="both", expand=True, padx=12, pady=8)

        frame_esq = ttk.Frame(corpo, padding=8)
        frame_dir = ttk.Frame(corpo, padding=8)
        corpo.add(frame_esq, weight=1)
        corpo.add(frame_dir, weight=2)

        ttk.Label(frame_esq, text="Arquivos selecionados", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.lista_arquivos = tk.Listbox(frame_esq, height=18)
        self.lista_arquivos.pack(fill="both", expand=True, pady=(8, 0))

        ttk.Label(frame_dir, text="Resultado / OBS", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.txt_resultado = tk.Text(frame_dir, wrap="word", font=("Consolas", 11), height=10)
        self.txt_resultado.pack(fill="both", expand=True, pady=(8, 8))

        ttk.Label(frame_dir, text="Log técnico", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.txt_log = tk.Text(frame_dir, wrap="word", font=("Consolas", 10), height=14)
        self.txt_log.pack(fill="both", expand=True)

        rodape = ttk.Frame(self.root, padding=12)
        rodape.pack(fill="x")
        ttk.Label(
            rodape,
            text="Dica: selecione apenas documentos do mesmo SEI para a OBS sair correta.",
            font=("Segoe UI", 9),
        ).pack(anchor="w")

    def selecionar_arquivos(self) -> None:
        caminhos = filedialog.askopenfilenames(
            title="Selecione os PDFs do SEI",
            filetypes=[("Arquivos PDF", "*.pdf")],
        )
        if not caminhos:
            return

        self.caminhos_arquivos = list(caminhos)
        self.lista_arquivos.delete(0, tk.END)
        for caminho in self.caminhos_arquivos:
            self.lista_arquivos.insert(tk.END, os.path.basename(caminho))

        self.label_status.config(text=f"{len(self.caminhos_arquivos)} arquivo(s) selecionado(s).")
        self._log("Arquivos selecionados com sucesso.")

    def limpar(self) -> None:
        self.caminhos_arquivos = []
        self.lista_arquivos.delete(0, tk.END)
        self.txt_resultado.delete("1.0", tk.END)
        self.txt_log.delete("1.0", tk.END)
        self.label_status.config(text="Nenhum arquivo selecionado.")

    def processar(self) -> None:
        if not self.caminhos_arquivos:
            messagebox.showwarning("Aviso", "Selecione ao menos um PDF.")
            return

        self.txt_resultado.delete("1.0", tk.END)
        self._log("Iniciando processamento...")
        self.label_status.config(text="Processando...")

        thread = threading.Thread(target=self._processar_em_thread, daemon=True)
        thread.start()

    def _processar_em_thread(self) -> None:
        try:
            documentos = carregar_documentos(self.caminhos_arquivos)
            analise = analisar_documentos(documentos)
            obs = gerar_obs_com_ia(texto_total)

            oficio = next(
                (
                    doc for doc in documentos
                    if doc.tipo_documento == "OFICIO" and "CEHAB/GOP" in doc.texto.upper()
                ),
                None
            )

            nome_dest, cargo_dest, destinatario_completo = (
                extrair_nome_e_cargo_destinatario(oficio.texto)
                if oficio else (None, None, None)
            )
            orgao_destino = normalizar_orgao_destino(destinatario_completo or "")

            detalhes = [
                "OBS GERADA:",
                obs,
                "",
                "RESUMO DA ANÁLISE:",
                f"- Data da solicitação GOP: {analise.data_solicitacao_gop or 'não identificada'}",
                f"- Há reiteração: {'sim' if analise.possui_reiteracao else 'não'}",
                f"- Dotação orçamentária: {'sim' if analise.dotacao_orcamentaria else 'não'}",
                f"- Em análise: {'sim' if analise.em_analise else 'não'}",
                f"- Programação financeira: {'sim' if analise.programacao_financeira else 'não'}",
                f"- Desdobramento de fonte: {'sim' if analise.desdobramento_fonte else 'não'}",
                f"- Autorização de execução: {'sim' if analise.autorizacao_execucao else 'não'}",
                f"- Destaque realizado: {'sim' if analise.destaque_realizado else 'não'}",
                f"- Órgão atual identificado: {analise.orgao_atual or 'não identificado'}",
                "",
                "DESTINATÁRIO DO OFÍCIO:",
                f"- Nome: {nome_dest or 'não identificado'}",
                f"- Cargo/órgão: {cargo_dest or 'não identificado'}",
                f"- Destinatário completo: {destinatario_completo or 'não identificado'}",
                f"- Sigla normalizada do órgão: {orgao_destino or 'não identificado'}",
                "",
                "DOCUMENTOS LIDOS:",
            ]

            for doc in documentos:
                detalhes.append(
                    f"- {doc.nome_arquivo} | tipo={doc.tipo_documento} | código={doc.codigo_documento or '-'} | data={doc.data_assinatura or '-'}"
                )

            self.root.after(0, self._mostrar_resultado, "\n".join(detalhes), obs)

        except Exception as e:
            self.root.after(0, self._erro_processamento, str(e))

        texto_total = "\n\n".join([doc.texto for doc in documentos])

    def _mostrar_resultado(self, detalhes: str, obs: str) -> None:
        self.txt_resultado.delete("1.0", tk.END)
        self.txt_resultado.insert("1.0", obs)
        self._log(detalhes)
        self.label_status.config(text="Processamento concluído.")

    def _erro_processamento(self, erro: str) -> None:
        self.label_status.config(text="Erro no processamento.")
        self._log(f"ERRO: {erro}")
        messagebox.showerror("Erro", erro)

    def copiar_obs(self) -> None:
        obs = self.txt_resultado.get("1.0", tk.END).strip()
        if not obs:
            messagebox.showinfo("Informação", "Nenhuma OBS gerada ainda.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(obs)
        self.root.update()
        messagebox.showinfo("Copiado", "OBS copiada para a área de transferência.")

    def _log(self, mensagem: str) -> None:
        self.txt_log.insert(tk.END, mensagem + "\n")
        self.txt_log.see(tk.END)


def gerar_obs_com_ia(texto: str) -> str:
    prompt = f"""
    Você é um especialista em análise de documentos do SEI (CEHAB/GOP).

    Gere uma OBS padrão com base no texto abaixo:

    - Identifique o ofício principal
    - Detecte reiteração (se houver)
    - Ordene cronologicamente
    - Gere no padrão formal

    Texto:
    {texto}
    """

    resposta = client.responses.create(
        model="gpt-4.1-mini",
        input=prompt,
        max_output_tokens=300
    )

    return resposta.output_text

def main() -> None:
    root = tk.Tk()
    app = AppSEIObs(root)
    root.mainloop()


if __name__ == "__main__":
    main()
