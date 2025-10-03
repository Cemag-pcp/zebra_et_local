import os
import json
import redis
import win32print

REDIS_URL = os.getenv("REDIS_URL", "redis://default:AWbmAbD4G2CfZPb3RxwuWQ4RfY7JOmxS@redis-19210.c262.us-east-1-3.ec2.redns.redis-cloud.com:19210")
QUEUE_NAME = os.getenv("REDIS_QUEUE", "print-zebra")  # nome da fila

def send_label(printer: str):
    """Fica escutando a fila no Redis e imprime cada job na impressora indicada."""
    r = redis.from_url(REDIS_URL)
    print(f"Conectado ao Redis. Aguardando jobs em '{QUEUE_NAME}' ? {printer}")

    while True:
        # BLPOP bloqueia até chegar um item
        _, raw = r.blpop(QUEUE_NAME)
        # raw é bytes; pode ser JSON {"zpl": "..."} ou ZPL puro
        try:
            payload = json.loads(raw)
            zpl = payload.get("zpl", "")
            if not zpl:
                print("[AVISO] JSON sem campo 'zpl'. Ignorando.")
                continue
        except Exception:
            # não é JSON; assume ZPL puro
            zpl = raw.decode("utf-8", errors="replace")

        try:
            hPrinter = win32print.OpenPrinter(printer)
            try:
                win32print.StartDocPrinter(hPrinter, 1, ("Etiqueta", None, "RAW"))
                win32print.StartPagePrinter(hPrinter)
                win32print.WritePrinter(hPrinter, zpl.encode("utf-8"))
                win32print.EndPagePrinter(hPrinter)
                win32print.EndDocPrinter(hPrinter)
                print("[OK] Etiqueta enviada para impressão.")
            finally:
                win32print.ClosePrinter(hPrinter)
        except Exception as e:
            print("[ERRO] Falha ao imprimir:", e)

if __name__ == "__main__":
    # ajuste o nome exato da sua impressora (veja em Dispositivos e Impressoras)
    send_label("ZDesigner ZD220-203dpi ZPL")
