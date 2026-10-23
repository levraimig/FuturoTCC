import os
import tempfile
import unittest
from datetime import datetime
from unittest.mock import MagicMock, patch
from pypdf import PdfWriter

from bot import LabimedBot


class FakeItem:
    def __init__(self, data_txt: str):
        self.data_txt = data_txt
        self.clicked = False

    def click(self):
        self.clicked = True

    def find_element(self, by, value):
        el = MagicMock()
        el.text = self.data_txt
        return el


class TestLabimedBot(unittest.TestCase):
    @patch("bot.time.sleep", return_value=None)
    @patch("bot.esperar_download")
    def test_fluxo_executar_fecha_popup_e_salva_pdfs(self, mock_esperar, _mock_sleep):
        with tempfile.TemporaryDirectory() as tmp:
            pasta_temp = os.path.join(tmp, "temp")
            os.makedirs(pasta_temp, exist_ok=True)
            pasta_internacao = os.path.join(tmp, "Paciente", "28-05 - internacao")

            logs = []
            bot = LabimedBot(
                "user123",
                "pass123",
                pasta_temp,
                logs.append,
                identificar_exames=True,
                unificar_pdfs=True,
            )

            bot.navegador = MagicMock()
            bot.navegador.current_url = "https://lis.labimed.com.br/shift/lis/labimed/elis/s01.iu.web.Login.cls#dadosOs"
            bot.wait = MagicMock()

            itens = [
                FakeItem("28/05/2025 - 10:00"),
                FakeItem("28/05/2025 - 14:30"),
                FakeItem("01/06/2025 - 09:00"),
            ]

            def side_effect_find_elements(by, val):
                if val == "listaOS":
                    return itens
                if val == "//input[@value='X']":
                    return [MagicMock()]
                return []

            bot.navegador.find_elements.side_effect = side_effect_find_elements

            def criar_pdf_falso(pasta, *args, **kwargs):
                caminho = os.path.join(pasta, "laudo_temp.pdf")
                w = PdfWriter()
                w.add_blank_page(width=200, height=200)
                with open(caminho, "wb") as f:
                    w.write(f)
                return "laudo_temp.pdf"

            mock_esperar.side_effect = criar_pdf_falso

            with patch.object(
                bot, "fechar_popup", wraps=bot.fechar_popup
            ) as spy_fechar:
                internacoes = {
                    pasta_internacao: (
                        datetime(2025, 5, 28, 0, 0),
                        datetime(2025, 5, 28, 23, 59),
                    )
                }
                ok, err = bot.executar(internacoes)

                self.assertEqual(ok, 2)
                self.assertEqual(err, 0)
                self.assertEqual(spy_fechar.call_count, 4)

            arquivos_gerados = sorted(os.listdir(pasta_internacao))
            self.assertIn("LAB1.pdf", arquivos_gerados)
            self.assertIn("LAB2.pdf", arquivos_gerados)
            self.assertIn("LAUDOS_UNIFICADOS.pdf", arquivos_gerados)

            bot.fechar_browser()
            self.assertFalse(os.path.exists(pasta_temp))

    def test_falha_login_levanta_excecao(self):
        with tempfile.TemporaryDirectory() as tmp:
            bot = LabimedBot("u", "s", tmp, lambda m: None)
            with patch.object(bot, "fazer_login", return_value=False):
                with self.assertRaises(RuntimeError):
                    bot.executar({})


if __name__ == "__main__":
    unittest.main()
