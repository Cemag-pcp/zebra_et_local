import win32print
import win32ui
# para converter a logo: curl --request POST http://api.labelary.com/v1/graphics --form file=@logo.png > logo.zpl

printer_name = "ZDesigner ZD220-203dpi ZPL"

logo_gfa = open('logo.zpl').read().strip()
logo_gfa_block = logo_gfa.replace("^FO0,0", "", 1)

zpl_code = f"""
^XA
^CI28
^PW800
^LL320
^LT0
^LH0,0

^FX Logo centralizado
^FO220,8{logo_gfa_block}

^FX Código da peça e descrição
^FO50,100^A0N,28,28^FDCódigo: 12345 - Peça Exemplo^FS

^FX Linha divisória vertical
^FO400,140^GB2,200,2^FS    ; linha de 2 dots de largura, 200 dots de altura

^FX Coluna PCP (esquerda)
^FO50,140^A0N,36,36^FB350,1,0,C,0^FDPCP^FS
^FO60,190^A0N,30,30^FDData: 31/07/2025^FS
^FO60,220^A0N,30,30^FDMáquina: MX-01^FS
^FO60,250^A0N,30,30^FDCambão: CB123^FS
^FO60,280^A0N,30,30^FDCor: Azul^FS

^FX Coluna Qualidade (direita)
^FO420,140^A0N,36,36^FB330,1,0,C,0^FDQualidade^FS
^FO430,200^A0N,30,30^FDAprovado:^FS
^FO560,200^A0N,30,30^FDSim^FS
^FO610,200^GB20,20,2^FS       ; caixa de 20×20 dots vazia
^FO650,200^A0N,30,30^FDNão^FS
^FO700,200^GB20,20,2^FS       ; caixa de 20×20 dots vazia
^FO430,230^A0N,30,30^FDData: ___/___/___^FS
^FO430,260^A0N,30,30^FDInspetor: _____________^FS

^XZ
"""

h = win32print.OpenPrinter(printer_name)
win32print.StartDocPrinter(h, 1, ("Etiqueta", None, "RAW"))
win32print.StartPagePrinter(h)
win32print.WritePrinter(h, zpl_code.encode("utf-8"))
win32print.EndPagePrinter(h)
win32print.EndDocPrinter(h)
win32print.ClosePrinter(h)
