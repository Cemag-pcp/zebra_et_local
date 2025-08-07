import streamlit as st
import pandas as pd
import win32print

# Função de envio de etiqueta já definida anteriormente
def send_label(printer, logo_gfa_block, file_name, qr_content):

    logo_gfa_block = logo_gfa_block.replace("^FO0,0", "", 1)

    # Monta ZPL
    zpl = f"""~DGR:logo.GRF,1234,567,:Z64:...seus_dados...  
^XA
^CI28
^PW800
^LL320
^LT0
^LH0,0

^FX Logo da empresa
^FO220,20{logo_gfa_block}

^FX Texto
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
    # Envia para a impressora
    hPrinter = win32print.OpenPrinter(printer)
    win32print.StartDocPrinter(hPrinter, 1, ("Etiqueta", None, "RAW"))
    win32print.StartPagePrinter(hPrinter)
    win32print.WritePrinter(hPrinter, zpl.encode('utf-8'))
    win32print.EndPagePrinter(hPrinter)
    win32print.EndDocPrinter(hPrinter)
    win32print.ClosePrinter(hPrinter)

# Streamlit App
def main():
    st.set_page_config(page_title="Label Printer", layout="wide")
    st.title("Impressão de Etiquetas Zebra ZD220")

    st.markdown("""
    Faça upload de um arquivo CSV com colunas:
    - **filename** (texto a exibir)
    - **url** (conteúdo do QR Code)
    """)

    uploaded_file = st.file_uploader("Escolha o arquivo CSV", type=["csv"])
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            df = df.iloc[0:1]

            if 'filename' in df.columns and 'url' in df.columns:
                st.success("CSV carregado com sucesso!")
                st.dataframe(df)

                if st.button("Imprimir etiquetas"):

                    logo_gfa = open('logo.zpl').read().strip()

                    with st.spinner("Enviando etiquetas para a impressora..."):
                        for _, row in df.iterrows():
                            send_label(
                                "ZDesigner ZD220-203dpi ZPL",
                                logo_gfa,
                                row['filename'],
                                row['url']
                            )
                        st.success("Todas as etiquetas foram enviadas!")
            else:
                st.error("O CSV deve conter as colunas 'filename' e 'url'.")
        except Exception as e:
            st.error(f"Erro ao ler o CSV: {e}")

if __name__ == "__main__":
    main()
