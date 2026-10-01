"""
Garante que os modulos do projeto (cpm_core, mix_processor_4, filtro_mix...)
sejam importaveis pelos testes em tests/, independente de onde o pytest for
chamado.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
