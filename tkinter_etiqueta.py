import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import pandas as pd
import os
import win32print
import time

# ============== ENVIO RAW (USB/Windows) ==============
def send_raw_windows(zpl: str, printer_name: str) -> int:
    data = zpl.encode("cp437", errors="replace")
    h = win32print.OpenPrinter(printer_name)
    try:
        win32print.StartDocPrinter(h, 1, ("RAW ZPL", None, "RAW"))
        win32print.StartPagePrinter(h)
        written = win32print.WritePrinter(h, data)
        win32print.EndPagePrinter(h)
        win32print.EndDocPrinter(h)
        return written or 0
    finally:
        win32print.ClosePrinter(h)

# ============== MODELOS (5) ==============
# Observação importante:
# - NÃO colocamos ~DGR placeholders. Vamos injetar o bloco do logo exatamente como vem do arquivo .zpl,
#   apenas removendo a 1ª ocorrência de "^FO0,0" (igual seu script).
# - Modelos que não usam logo podem ignorar logo_gfa_block.

def zpl_evandro():
    # Layout baseado no seu main_evandro.py (simplificado)
    return f"""^XA
^CI28
^PW800
^LL320
^LT0
^LH0,0

^FX Texto
^FO240,60
^A0N,35,35
^FB400,10,10,L,2
^FDProduto: Antirespingo de Solda Líquido Spatter Off^FS

^FO240,200
^A0N,35,35
^FB400,10,10,L,0
^FDCodigo: 350023^FS

^FO240,260
^A0N,35,35
^FB400,10,10,L,0
^FDConteúdo: 20L^FS

^XZ
"""

def zpl_laura(filename: str, url: str) -> str:
    filename = (filename or "").replace(".pdf", "")
    return f"""^XA
^CI28
^PW560
^LL240
^LT0
^LH0,0

^FX Logo
^FO50,10
^GFA,3800,3800,38,,::::::::::iP01F8,iP03FE,iP0607,iP0C03,iO018F18,iO01898C,iO018B0C,iO018F0C,iO018B0C,L03IFL03IF8N03FFI03FFCN03IFQ07FFEI01I9C,K03KFJ01KFM01IFE00JF8L03KFO07JFEI0C998,K0LFCI07KFEL07JF83JFEL0LFCM01LF800E03,J03MF003MF8J01KFC7KFK03MFM07LFE0070E,J0NFC07MFCJ03RF8J0NFCK01NF801FC,I01NFE1OFJ07RFCI01NFEK03NFC,I03OF3OF8I0SFEI03OF8J07NFE,I0OFE7OFC001TFI07OFCJ0PF8,001OFCPFE003TF800PFEI01PFC,001OF9QF003TF801QFI07PFE,003JF807FF3JFE07JF807IF8KFC3IFC03JFC07JFI07JF00JFE,007IFCI0FE7IFEI0JFC07FFE01JF00IFC07IFCI0JF800JF8001JF,00JFJ03C7IF8I03IFC0IF800IFE007FFC0JFJ03IFC01IFEJ07IF8,00IFCK08IFEK0IFE0IF8007FFC003FFE0IFCK0IFE01IF8J01IF8,01IF8L0IFCK07IF0IFI07FFC001FFE1IF8K07FFE03IFL0IFC,01IFL01IF8K03IF0IFI03FFC001FFE1IFL03IF03FFEL07FFC,03FFEL01IFL01IF8IFI03FFC001FFE3FFEL01IF07FFCL03FFE,03FFEL03FFEM0IF8IFI03FFC001FFE3FFEM0IF07FFCL01FFE,03FFCL03FFEM07FF8IFI03FFC001FFE7FFCM0IF8IF8L01IF,07FF8L03FFCM07FF8IFI03FFC001FFE7FF8M07FF8IFN0IF,07FF8L07SFCIFI03FFC001FFE7FF8M07FF8IFN0IF,07FF8L07SFCIFI03FFC001FFE7FF8M03FF8IFN07FF,07FFM07SFCIFI03FFC001FFE7FFN03FFCFFEN07FF8,:07FFM07SFCIFI03FFC001FFEIFN03FFDFFEN07FF8,:0IFM07SFCIFI03FFC001FFEIFN03FFDFFEN07FF8,07FFM07SFCIFI03FFC001FFE7FFN03FFCFFEN07FF8,:07FFM07SFCIFI03FFC001FFE7FF8M03FFCIFN07FF8,07FF8L07SFCIFI03FFC001FFE7FF8M07FFCIFN0IF8,07FF8L07FFCQ0IFI03FFC001FFE7FF8M07FFCIFN0IF8,07FFCL03FFCQ0IFI03FFC001FFE7FFCM0IFCIF8L01IF8,03FFCL03FFEQ0IFI03FFC001FFE3FFCM0IFC7FF8L01IF8,03FFEL03IFQ0IFI03FFC001FFE3FFEL01IFC7FFCL03IF8,01IFL01IF8P0IFI03FFC001FFE3IFL03IFC3FFEL07IF8,01IF8K01IFCP0IFI03FFC001FFE1IF8K07IFC3IFL0JF8,00IFCK08IFEK0EJ0IFI03FFC001FFE0IFCK0JFC1IF8J01JF8,00JFJ03CJFJ01F8I0IFI03FFC001FFE0JFJ03JFC1IFEJ07JF8,007IFCI07E7IFCI07FCI0IFI03FFC001FFE07IFCI07JFC0JF8I0KF8,003JF803FF3JFC03FFEI0IFI03FFC001FFE03JF003KFC07IFE007KF8,003OF9PFI0IFI03FFC001FFE03RFC07RF8,001OFCPF800IFI03FFC001FFE01RFC03RF8,I0OFE7OFC00IFI03FFC001FFE00RFC01RF8,I07OF3OF800IFI03FFC001FFE007QFC00RF8,I01NFE1OFI0IFI03FFC001FFE001QFC003QF8,J0NFC0NFEI0IFI03FFC001FFEI0QFC001QF8,J03MF003MF8I0IFI03FFC001FFEI03MF3FFCI07LFE7FF8,K0LFEI0LFEJ0IFI03FFC001FFEJ0LFE3FFCI01LFC7FF8,K03KFJ03KF8J0IFI03FFC001FFEJ03KF03FFCJ07JFE07FF8,L07IF8K07IFCK0IFI03FFC001FFEK07IF803FFCK0JF007FF,N02N02EM0402I01018I080CL033I01008L02J0IF,iN0IF,:iM01IF,hY01CL03FFE,hY07EL07FFE,hX01FFL07FFC,hX07FF8J01IFC,hX0IFCJ03IF8,hX0JFJ0JF8,hX0JFE003JF,hX07PFE,hX03PFC,hX01PF8,hY0PF,hY07NFE,hY01NF8,i0NF,i03LFC,iG0LF,iG01JF8,iH03FFC,,::::::^FS

^FX Texto principal
^FO35,105
^A0N,22,22
^FB320,3,10,L,0
^FD{filename}^FS

^FX Observação
^FO40,200
^A0N,17,17
^FB350,2,10,C,0
^FDPARA ASSISTENCIA TÉCNICA, ESCANEIE O QR CODE^FS

^FX QR
^FO385,60
^BQN,2,4
^FDHA,{url}^FS

^XZ
"""

def zpl_massari_simples(filename: str, url: str) -> str:
    filename = (filename or "").replace(".pdf", "")
    return f"""^XA
^PW800
^LL320
^LT0
^LH0,0

^FX Logo
^FO220,20
^GFA,3800,3800,38,,::::::::::iP01F8,iP03FE,iP0607,iP0C03,iO018F18,iO01898C,iO018B0C,iO018F0C,iO018B0C,L03IFL03IF8N03FFI03FFCN03IFQ07FFEI01I9C,K03KFJ01KFM01IFE00JF8L03KFO07JFEI0C998,K0LFCI07KFEL07JF83JFEL0LFCM01LF800E03,J03MF003MF8J01KFC7KFK03MFM07LFE0070E,J0NFC07MFCJ03RF8J0NFCK01NF801FC,I01NFE1OFJ07RFCI01NFEK03NFC,I03OF3OF8I0SFEI03OF8J07NFE,I0OFE7OFC001TFI07OFCJ0PF8,001OFCPFE003TF800PFEI01PFC,001OF9QF003TF801QFI07PFE,003JF807FF3JFE07JF807IF8KFC3IFC03JFC07JFI07JF00JFE,007IFCI0FE7IFEI0JFC07FFE01JF00IFC07IFCI0JF800JF8001JF,00JFJ03C7IF8I03IFC0IF800IFE007FFC0JFJ03IFC01IFEJ07IF8,00IFCK08IFEK0IFE0IF8007FFC003FFE0IFCK0IFE01IF8J01IF8,01IF8L0IFCK07IF0IFI07FFC001FFE1IF8K07FFE03IFL0IFC,01IFL01IF8K03IF0IFI03FFC001FFE1IFL03IF03FFEL07FFC,03FFEL01IFL01IF8IFI03FFC001FFE3FFEL01IF07FFCL03FFE,03FFEL03FFEM0IF8IFI03FFC001FFE3FFEM0IF07FFCL01FFE,03FFCL03FFEM07FF8IFI03FFC001FFE7FFCM0IF8IF8L01IF,07FF8L03FFCM07FF8IFI03FFC001FFE7FF8M07FF8IFN0IF,07FF8L07SFCIFI03FFC001FFE7FF8M07FF8IFN0IF,07FF8L07SFCIFI03FFC001FFE7FF8M03FF8IFN07FF,07FFM07SFCIFI03FFC001FFE7FFN03FFCFFEN07FF8,:07FFM07SFCIFI03FFC001FFEIFN03FFDFFEN07FF8,:0IFM07SFCIFI03FFC001FFEIFN03FFDFFEN07FF8,07FFM07SFCIFI03FFC001FFE7FFN03FFCFFEN07FF8,:07FFM07SFCIFI03FFC001FFE7FF8M03FFCIFN07FF8,07FF8L07SFCIFI03FFC001FFE7FF8M07FFCIFN0IF8,07FF8L07FFCQ0IFI03FFC001FFE7FF8M07FFCIFN0IF8,07FFCL03FFCQ0IFI03FFC001FFE7FFCM0IFCIF8L01IF8,03FFCL03FFEQ0IFI03FFC001FFE3FFCM0IFC7FF8L01IF8,03FFEL03IFQ0IFI03FFC001FFE3FFEL01IFC7FFCL03IF8,01IFL01IF8P0IFI03FFC001FFE3IFL03IFC3FFEL07IF8,01IF8K01IFCP0IFI03FFC001FFE1IF8K07IFC3IFL0JF8,00IFCK08IFEK0EJ0IFI03FFC001FFE0IFCK0JFC1IF8J01JF8,00JFJ03CJFJ01F8I0IFI03FFC001FFE0JFJ03JFC1IFEJ07JF8,007IFCI07E7IFCI07FCI0IFI03FFC001FFE07IFCI07JFC0JF8I0KF8,003JF803FF3JFC03FFEI0IFI03FFC001FFE03JF003KFC07IFE007KF8,003OF9PFI0IFI03FFC001FFE03RFC07RF8,001OFCPF800IFI03FFC001FFE01RFC03RF8,I0OFE7OFC00IFI03FFC001FFE00RFC01RF8,I07OF3OF800IFI03FFC001FFE007QFC00RF8,I01NFE1OFI0IFI03FFC001FFE001QFC003QF8,J0NFC0NFEI0IFI03FFC001FFEI0QFC001QF8,J03MF003MF8I0IFI03FFC001FFEI03MF3FFCI07LFE7FF8,K0LFEI0LFEJ0IFI03FFC001FFEJ0LFE3FFCI01LFC7FF8,K03KFJ03KF8J0IFI03FFC001FFEJ03KF03FFCJ07JFE07FF8,L07IF8K07IFCK0IFI03FFC001FFEK07IF803FFCK0JF007FF,N02N02EM0402I01018I080CL033I01008L02J0IF,iN0IF,:iM01IF,hY01CL03FFE,hY07EL07FFE,hX01FFL07FFC,hX07FF8J01IFC,hX0IFCJ03IF8,hX0JFJ0JF8,hX0JFE003JF,hX07PFE,hX03PFC,hX01PF8,hY0PF,hY07NFE,hY01NF8,i0NF,i03LFC,iG0LF,iG01JF8,iH03FFC,,::::::^FS

^FX Texto (nome)
^FO50,140
^A0N,30,30
^FB400,10,10,L,0
^FD{filename}^FS

^FX QR
^FO550,80
^BQN,2,4
^FDHA,{url}^FS

^XZ
"""

def zpl_massari_setor(filename: str, url: str, setor: str) -> str:
    filename = (filename or "").replace(".pdf", "")
    setor = setor or ""
    return f"""^XA
^CI28
^PW800
^LL320
^LT0
^LH0,0
^PR1
~SD12

^FX Logo
^FO220,20
^GFA,3800,3800,38,,::::::::::iP01F8,iP03FE,iP0607,iP0C03,iO018F18,iO01898C,iO018B0C,iO018F0C,iO018B0C,L03IFL03IF8N03FFI03FFCN03IFQ07FFEI01I9C,K03KFJ01KFM01IFE00JF8L03KFO07JFEI0C998,K0LFCI07KFEL07JF83JFEL0LFCM01LF800E03,J03MF003MF8J01KFC7KFK03MFM07LFE0070E,J0NFC07MFCJ03RF8J0NFCK01NF801FC,I01NFE1OFJ07RFCI01NFEK03NFC,I03OF3OF8I0SFEI03OF8J07NFE,I0OFE7OFC001TFI07OFCJ0PF8,001OFCPFE003TF800PFEI01PFC,001OF9QF003TF801QFI07PFE,003JF807FF3JFE07JF807IF8KFC3IFC03JFC07JFI07JF00JFE,007IFCI0FE7IFEI0JFC07FFE01JF00IFC07IFCI0JF800JF8001JF,00JFJ03C7IF8I03IFC0IF800IFE007FFC0JFJ03IFC01IFEJ07IF8,00IFCK08IFEK0IFE0IF8007FFC003FFE0IFCK0IFE01IF8J01IF8,01IF8L0IFCK07IF0IFI07FFC001FFE1IF8K07FFE03IFL0IFC,01IFL01IF8K03IF0IFI03FFC001FFE1IFL03IF03FFEL07FFC,03FFEL01IFL01IF8IFI03FFC001FFE3FFEL01IF07FFCL03FFE,03FFEL03FFEM0IF8IFI03FFC001FFE3FFEM0IF07FFCL01FFE,03FFCL03FFEM07FF8IFI03FFC001FFE7FFCM0IF8IF8L01IF,07FF8L03FFCM07FF8IFI03FFC001FFE7FF8M07FF8IFN0IF,07FF8L07SFCIFI03FFC001FFE7FF8M07FF8IFN0IF,07FF8L07SFCIFI03FFC001FFE7FF8M03FF8IFN07FF,07FFM07SFCIFI03FFC001FFE7FFN03FFCFFEN07FF8,:07FFM07SFCIFI03FFC001FFEIFN03FFDFFEN07FF8,:0IFM07SFCIFI03FFC001FFEIFN03FFDFFEN07FF8,07FFM07SFCIFI03FFC001FFE7FFN03FFCFFEN07FF8,:07FFM07SFCIFI03FFC001FFE7FF8M03FFCIFN07FF8,07FF8L07SFCIFI03FFC001FFE7FF8M07FFCIFN0IF8,07FF8L07FFCQ0IFI03FFC001FFE7FF8M07FFCIFN0IF8,07FFCL03FFCQ0IFI03FFC001FFE7FFCM0IFCIF8L01IF8,03FFCL03FFEQ0IFI03FFC001FFE3FFCM0IFC7FF8L01IF8,03FFEL03IFQ0IFI03FFC001FFE3FFEL01IFC7FFCL03IF8,01IFL01IF8P0IFI03FFC001FFE3IFL03IFC3FFEL07IF8,01IF8K01IFCP0IFI03FFC001FFE1IF8K07IFC3IFL0JF8,00IFCK08IFEK0EJ0IFI03FFC001FFE0IFCK0JFC1IF8J01JF8,00JFJ03CJFJ01F8I0IFI03FFC001FFE0JFJ03JFC1IFEJ07JF8,007IFCI07E7IFCI07FCI0IFI03FFC001FFE07IFCI07JFC0JF8I0KF8,003JF803FF3JFC03FFEI0IFI03FFC001FFE03JF003KFC07IFE007KF8,003OF9PFI0IFI03FFC001FFE03RFC07RF8,001OFCPF800IFI03FFC001FFE01RFC03RF8,I0OFE7OFC00IFI03FFC001FFE00RFC01RF8,I07OF3OF800IFI03FFC001FFE007QFC00RF8,I01NFE1OFI0IFI03FFC001FFE001QFC003QF8,J0NFC0NFEI0IFI03FFC001FFEI0QFC001QF8,J03MF003MF8I0IFI03FFC001FFEI03MF3FFCI07LFE7FF8,K0LFEI0LFEJ0IFI03FFC001FFEJ0LFE3FFCI01LFC7FF8,K03KFJ03KF8J0IFI03FFC001FFEJ03KF03FFCJ07JFE07FF8,L07IF8K07IFCK0IFI03FFC001FFEK07IF803FFCK0JF007FF,N02N02EM0402I01018I080CL033I01008L02J0IF,iN0IF,:iM01IF,hY01CL03FFE,hY07EL07FFE,hX01FFL07FFC,hX07FF8J01IFC,hX0IFCJ03IF8,hX0JFJ0JF8,hX0JFE003JF,hX07PFE,hX03PFC,hX01PF8,hY0PF,hY07NFE,hY01NF8,i0NF,i03LFC,iG0LF,iG01JF8,iH03FFC,,::::::^FS

^FX Texto (nome)
^FO50,140
^A0N,30,30
^FB400,10,10,L,0
^FD{filename}^FS

^FX SETOR
^FO200,250
^A0N,30,30
^FB400,1,10,C,0
^FD{setor}^FS

^FX QR
^FO550,80
^BQN,2,4
^FDHA,{url}^FS
^XZ
"""

def zpl_severiano_aprovado() -> str:
    # Modelo fixo do Severiano (sem CSV, já com ^GFA do logo embutido no arquivo original)
    return """^XA
^CI28
^PW560
^LL240
^LT0
^LH0,0
^FX Texto com código e nome (ajustado)
^FO200,30
^A0N,40,40
^FB320,1,10,L,0
^FDAPROVADO^FS

^FO140,80
^A0N,40,40
^FD ____/____/____^FS

^FO80,140
^A0N,40,40
^FD_____________________^FS

^FO395,185
^A0N,20,20
^FDResponsável^FS

^FX Logo da empresa
^FO50,26^GFA,8897,8897,41,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,iX07E,iX082,iW01B9,iW01A9,iW01A9,h0FFJ0FF8J01F807FK03FCL07FC01A9,gY0JF007IFJ0IF1FFEI03IFCJ03IF80C2,gX03FEDFC3FEDFC003FDIF7F800FFB7FI01FF6FF03C,gX07D7F7E7EBF7E007D77F5D7801F5FDF8001F5FBF,gW01F7D5D6F7EBD700F7DD5F7DC03DF575E003DF5EBC,gW01DD7F79DD7EFFC0FD77F5D7E0775FDFF00F75F7FE,gW03F783DBF7C3ABC1D7BD5F3D70FDF075700FDE0D56,gW075C00775C007EE3FC07F40FF1D7I0FD81D7001FF8,gW0FFJ07FI01BB3AC06B806B1FCI06FC3FCI06B8,gW0EBJ0EB8I0EF3FC03E807F1ACI03AC3ACI07E8,gW0FEJ0FEJ07BB54037806B3FCI01FC7F8I01BC,gV01D4J0D6J06EBFC03D807F75K0D66AJ01EC,gV01FCI01MFBB5C037806B7F8J0FE7FJ01BC,gV01ACI01D7JFDFBF403D807F6BK0D76AK0EE,gV01FCI01FD6DB6F5B5C037806B7FK0FDFEK0FA,gV01ACI01AKFBFBF403D807F6BK0D7AAK0DE,gV01FCI01FKAEAB5C037806B7FK0FD7EK0F6,gV01ACI01AKFBFBF403D807F6B8J0EF6BK0DE,gV01FCI01FKAEAB5C037806B7E8J0FB7FJ01F6,gV01AEJ0DEL03F403D807F6B8I01DF6B8I01DE,gW0FAJ0F7L035C037806B3FCI01F57E8I03F6,gW0DFJ0DD8K03F403D807F35CI03BF37CI075E,gW077800277C0030035C037806B1F7I07EB1DEI0FF6,gW07DF0077DF007C03F403D807F0DFC01EBF1F7803D5E,gW037F81FAF7C1FE03BC037806B0F5E07FEB0DFE07FF6,gW01D5IFDBDFFD703EC03D807F07F7FFABF0F5IF55E,gX0JF56EF7F7F037C037806B035FFEFEB03F7FDFF6,gX06AADFE7BD5D403D403D807F01F56BABF01BD6F55E,gX01IF501EFF78037C037806B007FFIEB006FFBEF6,gY06ABC007ABC003DC03D807FI0D578FF001AAE0DE,gY01FE8001FE8003F4037806BI07FC0D5I07F80F4,iV01DC,iV01F4,iQ018I03BC,iQ07CI07E8,iP01FFI0EB8,iP01EF803FE8,iQ0FBIFAA,iQ06EFFEFC,iQ03BADBAC,iQ01EFFEF,iR07AAB8,iR03FFE8,iS035,,,,,
^XZ
"""

def zpl_pernambuco_exportacao() -> str:
    # Modelo fixo do Pernambuco (sem CSV, já com ^GFA do logo embutido no arquivo original)
    return """^XA
^CI28
^PW560
^LL240
^LT0
^LH0,0

^FX Texto com código e nome (ajustado)
^FO120,80
^A0N,70,70
^FB400,1,10,L,0
^FDEXPORTAR^FS

^FX Logo da empresa
^FO50,26^GFA,8897,8897,41,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,iX07E,iX082,iW01B9,iW01A9,iW01A9,h0FFJ0FF8J01F807FK03FCL07FC01A9,gY0JF007IFJ0IF1FFEI03IFCJ03IF80C2,gX03FEDFC3FEDFC003FDIF7F800FFB7FI01FF6FF03C,gX07D7F7E7EBF7E007D77F5D7801F5FDF8001F5FBF,gW01F7D5D6F7EBD700F7DD5F7DC03DF575E003DF5EBC,gW01DD7F79DD7EFFC0FD77F5D7E0775FDFF00F75F7FE,gW03F783DBF7C3ABC1D7BD5F3D70FDF075700FDE0D56,gW075C00775C007EE3FC07F40FF1D7I0FD81D7001FF8,gW0FFJ07FI01BB3AC06B806B1FCI06FC3FCI06B8,gW0EBJ0EB8I0EF3FC03E807F1ACI03AC3ACI07E8,gW0FEJ0FEJ07BB54037806B3FCI01FC7F8I01BC,gV01D4J0D6J06EBFC03D807F75K0D66AJ01EC,gV01FCI01MFBB5C037806B7F8J0FE7FJ01BC,gV01ACI01D7JFDFBF403D807F6BK0D76AK0EE,gV01FCI01FD6DB6F5B5C037806B7FK0FDFEK0FA,gV01ACI01AKFBFBF403D807F6BK0D7AAK0DE,gV01FCI01FKAEAB5C037806B7FK0FD7EK0F6,gV01ACI01AKFBFBF403D807F6B8J0EF6BK0DE,gV01FCI01FKAEAB5C037806B7E8J0FB7FJ01F6,gV01AEJ0DEL03F403D807F6B8I01DF6B8I01DE,gW0FAJ0F7L035C037806B3FCI01F57E8I03F6,gW0DFJ0DD8K03F403D807F35CI03BF37CI075E,gW077800277C0030035C037806B1F7I07EB1DEI0FF6,gW07DF0077DF007C03F403D807F0DFC01EBF1F7803D5E,gW037F81FAF7C1FE03BC037806B0F5E07FEB0DFE07FF6,gW01D5IFDBDFFD703EC03D807F07F7FFABF0F5IF55E,gX0JF56EF7F7F037C037806B035FFEFEB03F7FDFF6,gX06AADFE7BD5D403D403D807F01F56BABF01BD6F55E,gX01IF501EFF78037C037806B007FFIEB006FFBEF6,gY06ABC007ABC003DC03D807FI0D578FF001AAE0DE,gY01FE8001FE8003F4037806BI07FC0D5I07F80F4,iV01DC,iV01F4,iQ018I03BC,iQ07CI07E8,iP01FFI0EB8,iP01EF803FE8,iQ0FBIFAA,iQ06EFFEFC,iQ03BADBAC,iQ01EFFEF,iR07AAB8,iR03FFE8,iS035,,,,,


^XZ
"""

def zpl_pernambuco_exportacao_V2() -> str:
    # Modelo fixo do Pernambuco (sem CSV, já com ^GFA do logo embutido no arquivo original)
    return """^XA
^CI28
^PW560
^LL240
^LT0
^LH0,0

^FX Texto com código e nome (ajustado)
^FO120,80
^A0N,70,70
^FB400,1,10,L,0
^FDEXPORTAR^FS

^FX Logo da empresa
^FO50,26^GFA,8897,8897,41,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,iX07E,iX082,iW01B9,iW01A9,iW01A9,h0FFJ0FF8J01F807FK03FCL07FC01A9,gY0JF007IFJ0IF1FFEI03IFCJ03IF80C2,gX03FEDFC3FEDFC003FDIF7F800FFB7FI01FF6FF03C,gX07D7F7E7EBF7E007D77F5D7801F5FDF8001F5FBF,gW01F7D5D6F7EBD700F7DD5F7DC03DF575E003DF5EBC,gW01DD7F79DD7EFFC0FD77F5D7E0775FDFF00F75F7FE,gW03F783DBF7C3ABC1D7BD5F3D70FDF075700FDE0D56,gW075C00775C007EE3FC07F40FF1D7I0FD81D7001FF8,gW0FFJ07FI01BB3AC06B806B1FCI06FC3FCI06B8,gW0EBJ0EB8I0EF3FC03E807F1ACI03AC3ACI07E8,gW0FEJ0FEJ07BB54037806B3FCI01FC7F8I01BC,gV01D4J0D6J06EBFC03D807F75K0D66AJ01EC,gV01FCI01MFBB5C037806B7F8J0FE7FJ01BC,gV01ACI01D7JFDFBF403D807F6BK0D76AK0EE,gV01FCI01FD6DB6F5B5C037806B7FK0FDFEK0FA,gV01ACI01AKFBFBF403D807F6BK0D7AAK0DE,gV01FCI01FKAEAB5C037806B7FK0FD7EK0F6,gV01ACI01AKFBFBF403D807F6B8J0EF6BK0DE,gV01FCI01FKAEAB5C037806B7E8J0FB7FJ01F6,gV01AEJ0DEL03F403D807F6B8I01DF6B8I01DE,gW0FAJ0F7L035C037806B3FCI01F57E8I03F6,gW0DFJ0DD8K03F403D807F35CI03BF37CI075E,gW077800277C0030035C037806B1F7I07EB1DEI0FF6,gW07DF0077DF007C03F403D807F0DFC01EBF1F7803D5E,gW037F81FAF7C1FE03BC037806B0F5E07FEB0DFE07FF6,gW01D5IFDBDFFD703EC03D807F07F7FFABF0F5IF55E,gX0JF56EF7F7F037C037806B035FFEFEB03F7FDFF6,gX06AADFE7BD5D403D403D807F01F56BABF01BD6F55E,gX01IF501EFF78037C037806B007FFIEB006FFBEF6,gY06ABC007ABC003DC03D807FI0D578FF001AAE0DE,gY01FE8001FE8003F4037806BI07FC0D5I07F80F4,iV01DC,iV01F4,iQ018I03BC,iQ07CI07E8,iP01FFI0EB8,iP01EF803FE8,iQ0FBIFAA,iQ06EFFEFC,iQ03BADBAC,iQ01EFFEF,iR07AAB8,iR03FFE8,iS035,,,,,


^XZ
"""

def zpl_pernambuco_cliente() -> str:
    # Modelo fixo do Pernambuco (sem CSV, já com ^GFA do logo embutido no arquivo original)
    return """^XA
^CI28
^PW800
^LL320
^LT0
^LH0,0

^FX Código da peça e descrição
^FO180,90^A0N,50,50^FDFORNECEDORA^FS

^FO180,190^A0N,50,50^FDMOSSORO-RN^FS

^XZ
"""

def zpl_matheus() -> str:
    # Modelo fixo do Pernambuco (sem CSV, já com ^GFA do logo embutido no arquivo original)
    return """^XA
^CI28
^PW800
^LL320
^LT0
^LH0,0

^FX Código da peça e descrição
^FO130,100^A0N,50,50^FDVAZAMENTO NO TERMINAL^FS

^XZ
"""

def zpl_teste_minimo() -> str:
    # ZPL mínimo p/ diagnosticar caminho spooler→Zebra
    return "^XA^FO50,50^ADN,36,20^FDTESTE ZPL OK^FS^XZ"

# ============== APP TKINTER ==============
class UnifiedUSBApp(tk.Tk):
    def __init__(self):

        super().__init__()
        self.title("Etiquetas Zebra – USB/Windows (5 modelos)")
        self.geometry("980x680")

        self.df = None
        self.logo_gfa = ""
        self.mode_requires_csv = False
        self.stop_event = threading.Event()
        self.printer_var = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        # Modelo
        ttk.Label(top, text="Finalidade / Modelo:").grid(row=0, column=0, sticky="w")
        self.model_var = tk.StringVar(value="Evandro (CSV filename/url)")
        self.model_combo = ttk.Combobox(
            top, textvariable=self.model_var, state="readonly",
            values=[
                "Evandro (CSV filename/url)",
                "Laura (CSV filename/url)",
                "Massari (Simples – CSV filename/url)",
                "Massari (com Setor – CSV filename/url/setor)",
                "Severiano – Aprovado (Fixa – sem CSV)",
                "Exportação - Pernambuco(Fixa – sem CSV)",
                "Cliente - Pernambuco(Fixa – sem CSV)",
                "Matheus (Fixa – sem CSV)",
            ], width=42
        )
        self.model_combo.grid(row=0, column=1, sticky="w", padx=(8,0))
        self.model_combo.bind("<<ComboboxSelected>>", self._on_model_change)

        # Quantidade
        ttk.Label(top, text="Quantidade:").grid(row=0, column=2, sticky="e", padx=(16,6))
        self.qtd_var = tk.StringVar(value="1")
        ttk.Entry(top, textvariable=self.qtd_var, width=10).grid(row=0, column=3, sticky="w")

        controls = ttk.Frame(self, padding=10)
        controls.pack(fill="x")
        ttk.Button(controls, text="Cancelar impressão (app)", command=self.on_cancel_app).pack(side="left")
        ttk.Button(controls, text="Parar agora na impressora (~JA)", command=self.on_cancel_printer).pack(side="left", padx=8)
        ttk.Button(controls, text="Resetar impressora (~JR)", command=self.on_reset_printer).pack(side="left", padx=8)
        ttk.Button(controls, text="Limpar RAM R: (^IDR:*.*)", command=self.on_clear_ram).pack(side="left", padx=8)

        # Arquivos / Ações
        mid = ttk.Frame(self, padding=(10,8,10,10))
        mid.pack(fill="x")

        self.btn_csv = ttk.Button(mid, text="Carregar CSV", command=self.load_csv)
        self.btn_csv.grid(row=0, column=1, sticky="w", padx=(8,0))

        self.csv_hint = ttk.Label(mid, text="", foreground="#555")
        self.csv_hint.grid(row=0, column=3, sticky="w", padx=(12,0))

        # Tabela CSV
        table = ttk.Frame(self, padding=(10,0,10,10))
        table.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(table, columns=("filename","url","setor"), show="headings", height=16)
        for c, w in (("filename", 260), ("url", 430), ("setor", 160)):
            self.tree.heading(c, text=c)
            self.tree.column(c, anchor="w", width=w)
        vsb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Progresso + Log + Botões
        bottom = ttk.Frame(self, padding=10)
        bottom.pack(fill="x")
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(fill="x", pady=(0,6))
        self.log = tk.Text(bottom, height=8, state="disabled")
        self.log.pack(fill="x")

        actions = ttk.Frame(self, padding=10)
        actions.pack(fill="x")
        ttk.Button(actions, text="Imprimir", command=self.on_print).pack(side="left")
        ttk.Button(actions, text="Imprimir teste (ZPL mínimo)", command=self.on_test).pack(side="left", padx=8)

        # Estado inicial
        self._on_model_change()

    def _on_model_change(self, *a):
        model = self.model_var.get()
        if model.startswith("Severiano") or model.startswith("Evandro") or model.startswith("Exportação") or model.startswith("Cliente") or model.startswith("Matheus"):
            self.mode_requires_csv = False
            self.csv_hint.config(text="Modelo FIXO. Não requer CSV. Quantidade = total.")
        elif "Setor" in model:
            self.mode_requires_csv = True
            self.csv_hint.config(text="CSV exigido: filename, url, setor. Quantidade = cópias por linha.")
        else:
            self.mode_requires_csv = True
            self.csv_hint.config(text="CSV exigido: filename, url. Quantidade = cópias por linha.")

    def _fill_table(self, df: pd.DataFrame):
        for i in self.tree.get_children():
            self.tree.delete(i)
        if df is None:
            return
        for _, r in df.iterrows():
            self.tree.insert("", "end", values=(r.get("filename",""), r.get("url",""), r.get("setor","")))

    def _log(self, msg: str):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def load_logo(self):
        path = filedialog.askopenfilename(title="Selecione logo.zpl", filetypes=[("ZPL", "*.zpl"), ("Todos", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                self.logo_gfa = f.read().strip()
            self._log(f"Logo carregado: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler logo.zpl:\n{e}")

    def load_csv(self):
        path = filedialog.askopenfilename(title="Selecione CSV", filetypes=[("CSV", "*.csv"), ("Todos", "*.*")])
        if not path:
            return
        try:
            df = pd.read_csv(path)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler CSV:\n{e}")
            return

        model = self.model_var.get()
        need = {"filename", "url"}
        if "Setor" in model:
            need = {"filename", "url", "setor"}

        if not need.issubset(df.columns):
            messagebox.showerror("Erro", f"CSV deve conter: {', '.join(sorted(need))}.")
            return

        self.df = df.copy()
        self._fill_table(self.df)
        self._log(f"CSV carregado ({len(self.df)} linhas).")

    def on_test(self):
        printer = "ZDesigner ZD220-203dpi ZPL"
        if not printer:
            messagebox.showwarning("Atenção", "Selecione uma impressora do Windows.")
            return
        zpl = zpl_teste_minimo()
        try:
            n = send_raw_windows(zpl, printer)
            self._log(f"[TESTE] Enviado {n} bytes para: {printer}")
        except win32print.error as e:
            self._log(f"[ERRO TESTE] win32: {e}")
            messagebox.showerror("Erro", f"Erro do Windows ao imprimir:\n{e}")
        except Exception as e:
            self._log(f"[ERRO TESTE] {e}")
            messagebox.showerror("Erro", f"Falha inesperada:\n{e}")

    def on_print(self):
        # valida impressora
        printer = "ZDesigner ZD220-203dpi ZPL"
        if not printer:
            messagebox.showwarning("Atenção", "Selecione uma impressora do Windows.")
            return

        # quantidade
        try:
            qtd = int((self.qtd_var.get() or "1").strip())
            assert qtd >= 1
        except Exception:
            messagebox.showwarning("Atenção", "Quantidade inválida (inteiro >= 1).")
            return

        # CSV se necessário
        if self.mode_requires_csv and (self.df is None or self.df.empty):
            messagebox.showwarning("Atenção", "Este modelo requer CSV carregado.")
            return

        model = self.model_var.get()

        # Thread para não travar UI
        t = threading.Thread(target=self._do_print, args=(printer, model, qtd), daemon=True)
        t.start()

    def _do_print(self, printer: str, model: str, qtd: int):
        self.progress["value"] = 0
        self.btn_state(False)
        self.stop_event.clear()
        try:
            total = qtd if not self.mode_requires_csv else len(self.df) * qtd
            self.progress["maximum"] = max(total, 1)
            done = 0

            def maybe_save(zpl: str):
                # salva ZPL para debug, se a flag existir e estiver True
                if getattr(self, "save_next_zpl", None) and self.save_next_zpl.get():
                    with open("last_print.zpl", "w", encoding="cp437", errors="replace") as f:
                        f.write(zpl)

            # --------- MODELO FIXO (sem CSV) ----------
            if not self.mode_requires_csv:
                for _ in range(qtd):
                    if self.stop_event.is_set():
                        break

                    # Apenas modelos realmente "fixos" entram aqui
                    if model.startswith("Severiano"):
                        zpl = zpl_severiano_aprovado()
                    elif model.startswith("Evandro"):
                        zpl = zpl_evandro()
                    elif model.startswith("Exportação"):
                        zpl = zpl_pernambuco_exportacao()
                    elif model.startswith("Cliente"):
                        zpl = zpl_pernambuco_cliente()
                    elif model.startswith("Matheus"):
                        zpl = zpl_matheus()
                    else:
                        self._log(f"[ERRO] O modelo '{model}' requer CSV. Selecione o CSV ou mude o modelo.")
                        break
                    
                    time.sleep(0.5)

                    maybe_save(zpl)
                    n = send_raw_windows(zpl, printer)
                    done += 1
                    self.progress["value"] = done
                    self._log(f"[OK] {model} | {n} bytes")

                if self.stop_event.is_set():
                    messagebox.showinfo("Cancelado", "Impressão cancelada pelo usuário.")
                else:
                    messagebox.showinfo("Sucesso", "Impressão finalizada.")
                return

            # --------- COM CSV ----------
            for _, row in self.df.iterrows():
                if self.stop_event.is_set():
                    break

                filename = str(row.get("filename", ""))
                url     = str(row.get("url", ""))
                setor   = str(row.get("setor", ""))

                for _ in range(qtd):
                    if self.stop_event.is_set():
                        break

                    # Seleção do modelo que usa CSV
                    if model.startswith("Laura"):
                        zpl = zpl_laura(filename, url)
                    elif "Setor" in model:
                        zpl = zpl_massari_setor(filename, url, setor)
                    else:
                        zpl = zpl_massari_simples(filename, url)
                    
                    time.sleep(0.5)

                    maybe_save(zpl)
                    n = send_raw_windows(zpl, printer)
                    done += 1
                    self.progress["value"] = done
                    self._log(f"[OK] {model} | {filename} | bytes={n}")

            if self.stop_event.is_set():
                messagebox.showinfo("Cancelado", "Impressão cancelada pelo usuário.")
            else:
                messagebox.showinfo("Sucesso", "Impressão finalizada.")

        except win32print.error as e:
            self._log(f"[ERRO] win32: {e}")
            messagebox.showerror("Erro", f"Erro do Windows ao imprimir:\n{e}")
        except Exception as e:
            self._log(f"[ERRO] {e}")
            messagebox.showerror("Erro", f"Falha inesperada:\n{e}")
        finally:
            self.btn_state(True)

    def btn_state(self, enabled: bool):
        # desabilita/habilita só o botão "Imprimir" para simplificar
        for w in self.children.values():
            pass
    
    def on_cancel_app(self):
        # Para o loop do app imediatamente
        self.stop_event.set()
        self._log("Cancelamento solicitado (app).")

    def on_cancel_printer(self):
        printer = "ZDesigner ZD220-203dpi ZPL"
        if not printer:
            messagebox.showwarning("Atenção", "Nenhuma impressora selecionada/instalada.")
            return
        try:
            send_raw_windows("~JA", printer)
            self._log("~JA enviado (cancelar formato atual).")
        except Exception as e:
            self._log(f"[ERRO ~JA] {e}")

    def on_reset_printer(self):
        printer = "ZDesigner ZD220-203dpi ZPL"
        if not printer:
            messagebox.showwarning("Atenção", "Selecione uma impressora.")
            return
        try:
            send_raw_windows("~JR", printer)
            self._log("~JR enviado (reset).")
        except Exception as e:
            self._log(f"[ERRO ~JR] {e}")

    def on_clear_ram(self):
        if not messagebox.askyesno("Confirmação", "Apagar TODOS os arquivos da RAM (R:)?"):
            return
        printer = "ZDesigner ZD220-203dpi ZPL"
        if not printer:
            messagebox.showwarning("Atenção", "Selecione uma impressora.")
            return
        try:
            zpl = "^XA^IDR:*.*^XZ"
            send_raw_windows(zpl, printer)
            self._log("RAM (R:) limpa com ^IDR:*.*.")
        except Exception as e:
            self._log(f"[ERRO limpar R:] {e}")

    def on_clear_flash(self):
        if not messagebox.askyesno("PERIGO", "Apagar TODOS os arquivos da FLASH (E:)? Isso remove logos/fonte salvos."):
            return
        printer = "ZDesigner ZD220-203dpi ZPL"
        if not printer:
            messagebox.showwarning("Atenção", "Selecione uma impressora.")
            return
        try:
            zpl = "^XA^IDE:*.*^XZ"
            send_raw_windows(zpl, printer)
            self._log("FLASH (E:) limpa com ^IDE:*.*.")
        except Exception as e:
            self._log(f"[ERRO limpar E:] {e}")

    def _get_selected_printer(self) -> str:
        """Retorna a impressora selecionada ou a default do Windows.
        Garante que self.printer_var exista."""
        # cria se não existir
        if not hasattr(self, "printer_var"):
            try:
                import win32print
                default_prn = win32print.GetDefaultPrinter()
            except Exception:
                default_prn = ""
            self.printer_var = tk.StringVar(value=default_prn)

        # se houver combobox, tenta ler; senão, usa o valor atual
        try:
            prn = self.printer_var.get().strip()
        except Exception:
            prn = ""

        if not prn:
            try:
                import win32print
                prn = win32print.GetDefaultPrinter()
            except Exception:
                prn = ""
        return prn or ""

if __name__ == "__main__":
    app = UnifiedUSBApp()
    app.mainloop()
