import win32print

def send_label(printer, logo_gfa_block, file_name, qr_content):
    """
    Gera e envia etiqueta com:
     - Logo (inline GFA)
     - Texto com código e nome de arquivo
     - QR Code com a URL, redimensionado para caber
    """
    # Posicionamento
    x_logo, y_logo = 50, 10
    x_text, y_text1, y_text2 = 50, 140, 180
    
    # QR Code reduzido e afastado da borda
    # module_size 4 (menor) e correção H
    qr_module_size = 4
    qr_ec_level   = 'H'
    # posição X ajustada para não ficar fora da largura
    x_qr, y_qr    = 550, 80  
    
    logo_gfa_block = logo_gfa_block.replace("^FO0,0", "", 1)

    # Monta ZPL
    zpl = f"""^XA
^PW800
^LL320
^LT0
^LH0,0

^FX Logo da empresa
^FO220,20{logo_gfa_block}

^FX Texto com código e nome, quebrando linha se necessário
^FO50,140
^A0N,30,30        ; fonte 30×30 dots
^FB400,10,10,L,0   ; largura=400, max=2 linhas, 10 dots de espaçamento, alinhamento Left, sem indent
^FD{file_name.replace(".pdf","")}^FS

^FX QR Code
^FO550,80
^BQN,2,4
^FDHA,{qr_content}^FS

^XZ
"""

    hPrinter = win32print.OpenPrinter(printer)
    win32print.StartDocPrinter(hPrinter, 1, ("Etiqueta", None, "RAW"))
    win32print.StartPagePrinter(hPrinter)
    win32print.WritePrinter(hPrinter, zpl.encode("utf-8"))
    win32print.EndPagePrinter(hPrinter)
    win32print.EndDocPrinter(hPrinter)
    win32print.ClosePrinter(hPrinter)

# Exemplo de uso
import pandas as pd

df = pd.read_csv('labels.csv')
df = df.iloc[0:1]

logo_gfa = open('logo.zpl').read().strip()

for _, row in df.iterrows():
    send_label(
        "ZDesigner ZD220-203dpi ZPL",
        logo_gfa,
        row['filename'],
        row['url']
    )
