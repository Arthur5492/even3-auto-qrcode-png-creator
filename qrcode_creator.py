import qrcode
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os
from pathlib import Path
import re
import argparse
import sys
from datetime import datetime

# =========================
# Configurações principais
# =========================
# DPI alvo para as páginas A4 (mantendo consistência com as imagens individuais)
DPI = 150

# Tamanho A4 em pixels para o DPI escolhido: 8.27in x 11.69in
A4_W = int(8.27 * DPI)   # 1240 @150dpi
A4_H = int(11.69 * DPI)  # 1754 @150dpi

# Margens e espaçamentos (em pixels) – reduzidos ao mínimo seguro
MARGEM_X = 8
MARGEM_Y = 8
GUTTER_X = 8
GUTTER_Y = 8

# Tamanho do "crachá" individual (mesmo usado antes)
ITEM_W = 400
ITEM_H = 150


def _carregar_fonte(font_size=24):
    """Tenta carregar uma fonte boa; cai no default se não achar."""
    try:
        return ImageFont.truetype("arial.ttf", font_size)
    except (OSError, IOError):
        for path in (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/System/Library/Fonts/Helvetica.ttc",
        ):
            try:
                return ImageFont.truetype(path, font_size)
            except (OSError, IOError):
                continue
    return ImageFont.load_default()


def sanitize_codigo(valor):
    """
    Converte o código de inscrição para string, removendo '.0' quando vier de float,
    e garantindo apenas dígitos se fizer sentido.
    """
    if pd.isna(valor):
        return ""
    if isinstance(valor, (int,)):
        return str(valor)
    if isinstance(valor, float):
        if float(int(valor)) == float(valor):
            return str(int(valor))
        else:
            return str(valor).replace(",", ".")
    s = str(valor).strip()
    if re.match(r"^\d+\.0$", s):
        s = s[:-2]
    s = s.strip()
    return s


def gerar_qrcode_individual_img(nome_completo, codigo_inscricao):
    """
    Gera e retorna uma PIL.Image com QR code (à esquerda) e nome (à direita).
    NÃO salva em disco; retorna apenas o objeto de imagem.
    """
    cor_fundo = (255, 255, 255)
    cor_texto = (0, 0, 0)

    # QRCode
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=6,
        border=2,
    )
    qr.add_data(str(codigo_inscricao))
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white")
    qr_img = qr_img.resize((120, 120), Image.Resampling.LANCZOS)

    # Canvas do item
    img = Image.new("RGB", (ITEM_W, ITEM_H), cor_fundo)
    draw = ImageDraw.Draw(img)

    # QR à esquerda centralizado verticalmente
    qr_x = 15
    qr_y = (ITEM_H - 120) // 2
    img.paste(qr_img, (qr_x, qr_y))

    # Texto
    font = _carregar_fonte(24)

    texto_x_inicio = qr_x + 120 + 20
    texto_largura = ITEM_W - texto_x_inicio - 15

    # Quebra de linha manual para caber no retângulo
    palavras = str(nome_completo).split()
    linhas = []
    linha_atual = ""

    for palavra in palavras:
        teste = (linha_atual + " " + palavra).strip()
        bbox = draw.textbbox((0, 0), teste, font=font)
        largura_teste = bbox[2] - bbox[0]
        if largura_teste <= texto_largura:
            linha_atual = teste
        else:
            if linha_atual:
                linhas.append(linha_atual)
                linha_atual = palavra
            else:
                linhas.append(palavra)  # palavra enorme
                linha_atual = ""

    if linha_atual:
        linhas.append(linha_atual)

    # Centraliza verticalmente
    bbox_ag = draw.textbbox((0, 0), "Ag", font=font)
    altura_linha = bbox_ag[3] - bbox_ag[1]
    espacamento_linha = 5
    altura_total_texto = len(linhas) * altura_linha + (len(linhas) - 1) * espacamento_linha
    texto_y_inicio = (ITEM_H - altura_total_texto) // 2

    for i, linha in enumerate(linhas):
        y_linha = texto_y_inicio + i * (altura_linha + espacamento_linha)
        draw.text((texto_x_inicio, y_linha), linha, fill=cor_texto, font=font)

    return img


def montar_paginas_a4(imagens_tuplas, diretorio_saida, dpi=DPI):
    """
    Recebe uma lista de tuplas (PIL.Image, idx_linha_1based) e monta páginas A4 em grid,
    salvando SOMENTE as páginas A4 (PNG) com o nome '1.png', '2.png', ...
    Retorna lista de caminhos salvos.
    """
    if not imagens_tuplas:
        print("⚠️  Nenhuma imagem para montar nas páginas A4.")
        return []

    Path(diretorio_saida).mkdir(parents=True, exist_ok=True)

    # Dimensões úteis (área interna após margens)
    area_w = A4_W - 2 * MARGEM_X
    area_h = A4_H - 2 * MARGEM_Y

    # Quantidade de colunas/linhas que cabem
    cols = max(1, (area_w + GUTTER_X) // (ITEM_W + GUTTER_X))
    rows = max(1, (area_h + GUTTER_Y) // (ITEM_H + GUTTER_Y))
    por_pagina = int(cols * rows)

    # Centralização da grade
    grid_w = cols * ITEM_W + (cols - 1) * GUTTER_X
    grid_h = rows * ITEM_H + (rows - 1) * GUTTER_Y
    offset_x = MARGEM_X + (area_w - grid_w) // 2
    offset_y = MARGEM_Y + (area_h - grid_h) // 2

    salvos = []
    i = 0
    total = len(imagens_tuplas)
    page_num = 1  # << numerador de páginas

    while i < total:
        canvas = Image.new("RGB", (A4_W, A4_H), (255, 255, 255))

        for r in range(rows):
            for c in range(cols):
                if i >= total:
                    break
                img, _idx_linha = imagens_tuplas[i]
                x = offset_x + c * (ITEM_W + GUTTER_X)
                y = offset_y + r * (ITEM_H + GUTTER_Y)
                canvas.paste(img, (x, y))
                i += 1
            if i >= total:
                break

        nome_arq = f"{page_num}.png"
        caminho = os.path.join(diretorio_saida, nome_arq)
        canvas.save(caminho, "PNG", dpi=(dpi, dpi))
        salvos.append(caminho)
        print(f"🖨️  Página {page_num} salva: {caminho}  (itens por página: {por_pagina}, cols={cols}, rows={rows})")
        page_num += 1

    print(f"📄 Total de páginas A4: {len(salvos)}")
    return salvos


def juntar_paginas_em_pdf(caminhos_png, diretorio_saida, nome_pdf="badges_A4_all.pdf", dpi=DPI):
    """
    Recebe a lista de caminhos das páginas A4 (PNG) e salva todas em um único PDF.
    A ordem segue a lista recebida.
    """
    if not caminhos_png:
        print("⚠️  Nenhuma página PNG para juntar no PDF.")
        return None

    imagens = []
    for p in caminhos_png:
        try:
            img = Image.open(p).convert("RGB")
            imagens.append(img)
        except Exception as e:
            print(f"⚠️  Não foi possível abrir '{p}' para o PDF: {e}")

    if not imagens:
        print("⚠️  Nenhuma imagem válida para PDF.")
        return None

    pdf_path = os.path.join(diretorio_saida, nome_pdf)
    primeira, restantes = imagens[0], imagens[1:]
    primeira.save(pdf_path, "PDF", resolution=dpi, save_all=True, append_images=restantes)
    print(f"📚 PDF gerado com {len(imagens)} páginas: {pdf_path}")
    return pdf_path


def processar_planilha(caminho_xlsx, diretorio_saida, abrir_pasta=False):
    """
    Lê a planilha, ordena por 'Nome' crescente, gera TODAS as imagens individuais em memória
    e monta páginas A4. Salva apenas as páginas A4, nomeando como 'firstIdx-lastIdx.png'
    e, por fim, gera um PDF único com todas as páginas.
    
    Args:
        caminho_xlsx: Caminho do arquivo Excel
        diretorio_saida: Diretório onde salvar os arquivos
        abrir_pasta: Se True, abre a pasta automaticamente após gerar
    """
    print("=" * 60)
    print("🎯 GERADOR DE PÁGINAS A4 COM QR CODES (SEM SALVAR INDIVIDUAIS)")
    print("=" * 60)

    print(f"📖 Lendo planilha: {caminho_xlsx}")
    try:
        # força 'Número de Inscrição' como string para evitar '.0'
        df = pd.read_excel(
            caminho_xlsx,
            engine="openpyxl",                     # <- lê .xlsx e .xlsm
            dtype={'Número de Inscrição': 'string'}
        )
    except Exception as e:
        print(f"⚠️ Falha ao ler planilha com dtype forçado: {e}")
        df = pd.read_excel(caminho_xlsx, engine="openpyxl")

    # Verificações
    if 'Nome' not in df.columns:
        print("❌ ERRO: Coluna 'Nome' não encontrada na planilha!")
        return
    if 'Número de Inscrição' not in df.columns:
        print("❌ ERRO: Coluna 'Número de Inscrição' não encontrada na planilha!")
        return

    df['Data Inscrição'] = pd.to_datetime(
        df['Data Inscrição'].astype(str).str.strip(),
        format='%d/%m/%Y %H:%M:%S',   # ou use dayfirst=True se o formato variar
        errors='coerce'
    )
    
    imagens_tuplas = []  # [(PIL.Image, idx_linha_1based)]
    erros = 0
    sucessos = 0

    for idx, row in df.iterrows():
        nome = str(row.get('Nome', '')).strip()
        codigo = sanitize_codigo(row.get('Número de Inscrição', ''))
        idx_linha = idx + 1  # 1-based após ordenação

        if not nome or not codigo:
            print(f"⚠️  Linha {idx_linha}: nome/código ausente. Nome='{nome}', Código='{codigo}'")
            erros += 1
            continue

        try:
            img = gerar_qrcode_individual_img(nome, codigo)
            imagens_tuplas.append((img, idx_linha))
            sucessos += 1
        except Exception as e:
            print(f"❌ ERRO ao gerar item da linha {idx_linha}: {e}")
            erros += 1

    print(f"✅ Itens gerados: {sucessos}  |  ❌ Erros: {erros}")

    if not imagens_tuplas:
        print("⚠️ Nada para montar em A4.")
        return

    # Monta e salva páginas A4
    Path(diretorio_saida).mkdir(parents=True, exist_ok=True)
    caminhos_png = montar_paginas_a4(imagens_tuplas, diretorio_saida, dpi=DPI)

    # Junta tudo em um PDF
    if caminhos_png:
        juntar_paginas_em_pdf(caminhos_png, diretorio_saida, nome_pdf="badges_A4_all.pdf", dpi=DPI)

    
    # Tentar abrir pasta no SO
    if abrir_pasta:
        abs_path = os.path.abspath(diretorio_saida)
        try:
            os.path
            import subprocess, platform
            if platform.system() == "Windows":
                subprocess.run(['explorer', abs_path])
            elif platform.system() == "Darwin":
                subprocess.run(['open', abs_path])
            elif platform.system() == "Linux":
                subprocess.run(['xdg-open', abs_path])
        except Exception as e:
            print(f"⚠️ Não foi possível abrir a pasta automaticamente: {e}")



def parse_arguments():
    """
    Configura e processa os argumentos da linha de comando.
    """
    
    script_name = os.path.basename(sys.argv[0])
    
    parser = argparse.ArgumentParser(
        description='Gerador de QR Codes em páginas A4 a partir de planilha Excel',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f'''
Exemplos de uso:
  python {script_name} -file planilha.xlsx -output ./saida
  python {script_name} -file "C:\\Users\\nome\\planilha.xlsm" -output "C:\\Output"
  python {script_name} --help

O script lê uma planilha Excel (.xlsx ou .xlsm) com as colunas:
  - Nome: nome completo do participante
  - Número de Inscrição: código único para o QR code

Gera páginas A4 em PNG com os QR codes organizados em grid e um PDF final.
        '''
    )
    
    parser.add_argument(
        '-file',
        '--file',
        '-f',
        type=str,
        required=True,
        help='Caminho para o arquivo da planilha Excel (.xlsx ou .xlsm)'
    )
    
    parser.add_argument(
        '-output',
        '--output',
        '-o',
        type=str,
        default='./output',
        help='Caminho do diretório de saída para os arquivos gerados (padrão: ./output)'
    )
    
    parser.add_argument(
        '--open',
        action='store_true',
        help='Abre automaticamente a pasta de saída após gerar os arquivos'
    )
    
    return parser.parse_args()


def main():
    # Processa argumentos da linha de comando
    args = parse_arguments()
    
    caminho_xlsx = args.file
    diretorio_base = args.output
    abrir_pasta = args.open
    
    # Cria diretório com timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    diretorio_saida = os.path.join(diretorio_base, timestamp)
    
    print(f"📁 Arquivo de entrada: {caminho_xlsx}")
    print(f"📂 Diretório de saída: {diretorio_saida}")
    
    # Verifica se o arquivo existe
    if not os.path.exists(caminho_xlsx):
        print(f"❌ ERRO: Arquivo não encontrado:\n{caminho_xlsx}")
        sys.exit(1)
    
    # Processa a planilha
    processar_planilha(caminho_xlsx, diretorio_saida,abrir_pasta)


if __name__ == "__main__":
    main()