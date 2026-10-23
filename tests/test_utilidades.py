import os
import tempfile
import unittest
from datetime import datetime
from pypdf import PdfWriter

from utilidades import (
    converter_data_site,
    esperar_download,
    extrair_resumo_laudo_pdf,
    formatar_mascara_data_hora,
    montar_caminho_internacao,
    preparar_pastas,
    sanitizar_nome_arquivo,
    unificar_pdfs_internacao,
)


class TestUtilidades(unittest.TestCase):
    def test_converter_data_site(self):
        dt = converter_data_site("28/05/2025 - 13:50")
        self.assertEqual(dt, datetime(2025, 5, 28, 13, 50))
        self.assertIsNone(converter_data_site("data invalida"))
        self.assertIsNone(converter_data_site(""))

    def test_formatar_mascara_data_hora(self):
        self.assertEqual(
            formatar_mascara_data_hora("280520251350"), "28/05/2025 13:50"
        )
        self.assertEqual(formatar_mascara_data_hora("2805"), "28/05")
        self.assertEqual(formatar_mascara_data_hora("28/05/2025"), "28/05/2025")

    def test_sanitizar_nome_arquivo_e_reservados(self):
        self.assertEqual(
            sanitizar_nome_arquivo('João: Silva / Teste*?"<>|'), "João Silva Teste"
        )
        self.assertEqual(sanitizar_nome_arquivo("CON"), "Paciente_Sem_Nome")
        self.assertEqual(sanitizar_nome_arquivo("   "), "Paciente_Sem_Nome")

    def test_preparar_e_montar_pastas(self):
        with tempfile.TemporaryDirectory() as tmp:
            pasta_pac, pasta_temp = preparar_pastas("Maria Souza", tmp)
            self.assertTrue(os.path.isdir(pasta_pac))
            self.assertTrue(os.path.isdir(pasta_temp))

            usados = set()
            c1 = montar_caminho_internacao(
                pasta_pac, datetime(2025, 5, 10, 8, 0), usados
            )
            c2 = montar_caminho_internacao(
                pasta_pac, datetime(2025, 5, 10, 14, 0), usados
            )
            self.assertTrue(os.path.isdir(c1))
            self.assertTrue(os.path.isdir(c2))
            self.assertNotEqual(c1, c2)
            self.assertTrue(c2.endswith("10-05 - internacao (2)"))

    def test_esperar_download_ignora_crdownload_e_respeita_parada(self):
        with tempfile.TemporaryDirectory() as tmp:
            cr = os.path.join(tmp, "exame.pdf.crdownload")
            pdf = os.path.join(tmp, "exame.pdf")
            with open(cr, "wb") as f:
                f.write(b"parcial")
            w = PdfWriter()
            w.add_blank_page(width=200, height=200)
            with open(pdf, "wb") as f:
                w.write(f)

            self.assertIsNone(esperar_download(tmp, timeout=1, intervalo=0.2))
            self.assertIsNone(
                esperar_download(tmp, timeout=5, intervalo=0.2, deve_parar=lambda: True)
            )

            os.remove(cr)
            self.assertEqual(
                esperar_download(tmp, timeout=1, intervalo=0.2), "exame.pdf"
            )

    def test_unificar_pdfs_internacao(self):
        with tempfile.TemporaryDirectory() as tmp:
            for idx in (1, 2):
                writer = PdfWriter()
                writer.add_blank_page(width=200, height=200)
                with open(os.path.join(tmp, f"LAB{idx}.pdf"), "wb") as f:
                    writer.write(f)

            unificado = unificar_pdfs_internacao(tmp)
            self.assertIsNotNone(unificado)
            self.assertTrue(os.path.isfile(unificado))
            self.assertEqual(extrair_resumo_laudo_pdf(unificado), "")


if __name__ == "__main__":
    unittest.main()
