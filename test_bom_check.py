"""
Prueba aislada: verificar clases BOM de ZFER vs ZPLA
ZFER: 700060883  /  ZPLA: 503001496

Requiere SAP GUI abierto con sesión activa en QUAS.
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from sap_auto import AutomatizadorSAP

ZFER = "700163297"
ZPLA = "503001496"

class FakeRes:
    """Simula ResultadoCombinacion solo para capturar logs."""
    def __init__(self):
        self.advertencias = []
        self.logs = []

    def _advertir(self, msg):
        self.advertencias.append(msg)
        print(f"  [FAKE_RES._advertir] {msg}")

    def _log(self, msg):
        self.logs.append(msg)
        print(f"  [FAKE_RES._log] {msg}")


def main():
    print("=" * 60)
    print(f"TEST BOM-CHECK  ZFER={ZFER}  ZPLA={ZPLA}")
    print("=" * 60)

    auto = AutomatizadorSAP()
    print("\n[0] Conectando a SAP GUI...")
    if not auto.conectar():
        print("  ERROR: No se pudo conectar a SAP. Asegúrate de tener SAP GUI abierto con sesión activa.")
        sys.exit(1)
    print("  → Conectado OK")

    # 1) Leer BOM del ZFER desde ZPPR0008
    print(f"\n[1] Leyendo BOM del ZFER {ZFER} en ZPPR0008...")
    bom_raw = auto.zppr0008_leer_bom_completo(ZFER)
    filas   = bom_raw.get("filas", [])
    print(f"    → {len(filas)} filas leídas")
    for f in filas:
        print(f"      pos={f.get('pos')}  clase={f.get('clase')!r}  postp={f.get('postp')}")

    bom_zfer = [{"pos": f["pos"], "clase": f.get("clase", "")} for f in filas]

    # 2) Leer clases del ZPLA desde ZPPR0008 (sesión auxiliar)
    print(f"\n[2] Leyendo clases del ZPLA {ZPLA} en ZPPR0008...")
    clases = auto._leer_clases_zpla_sap(ZPLA)
    print(f"    → {len(clases)//2} posiciones (dict tiene entradas con y sin zero-pad)")
    for k, v in sorted(clases.items()):
        if len(k) == 4:  # mostrar solo los zero-padded para no duplicar
            print(f"      pos={k}  clase={v!r}")

    # 3) Comparar
    print(f"\n[3] Comparando clases ZFER vs ZPLA...")
    res = FakeRes()
    auto._verificar_clases_bom(bom_zfer, clases, ZPLA, res)

    # 4) Resumen
    print("\n" + "=" * 60)
    print("RESULTADO:")
    if res.advertencias:
        print(f"  ⚠  {len(res.advertencias)} advertencia(s):")
        for a in res.advertencias:
            print(f"     • {a}")
    else:
        print("  ✓  Sin diferencias de clase en posiciones coincidentes")
    print("=" * 60)


if __name__ == "__main__":
    main()
