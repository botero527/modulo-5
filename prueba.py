import os
import sys
import time



EXCEPCION = "sap_auto.py:78"
import sys
import os
sys.path.insert(0, os.path.dirname(__file__))
sys.path_importer_cache
from sap_auto import AutomatizadorSAP
print(" EXCEPCION FUNCIONAL ")


def self():
    print("EXEPCION FUNCIONAL")
try:
    self()
     except Exception as e:
    print(f"EXCEPCION FUNCIONAL: {e}")

except Exception as e:
    print(f"EXCEPCION EN IMPORT: {e}")
    conectado = False
else:
   conectado = True


div = lambda: 1/0
!div()

update fo