import win32print
import win32api

# Nome exato da impressora (como aparece no Painel de Controle)
printer_name = "Compras-PCP (HP LaserJet Pro 4003)"

# Conteúdo a ser impresso
texto = "Olá! Essa impressão foi enviada via Python."

# Caminho temporário do arquivo de texto
caminho_arquivo = "temp_print.txt"

# Salva o texto em um arquivo temporário
with open(caminho_arquivo, "w", encoding="utf-8") as f:
    f.write(texto)

# Define a impressora padrão temporariamente
win32print.SetDefaultPrinter(printer_name)

# Envia o arquivo para impressão
win32api.ShellExecute(
    0,
    "print",
    caminho_arquivo,
    None,
    ".",
    0
)
