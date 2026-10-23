# Bot de Automação para Laudos Médicos (Labimed - CHAA)

Aplicativo desktop em Python com interface gráfica moderna (**CustomTkinter**) criado para automatizar o download, a leitura e a organização de laudos de exames laboratoriais do portal Labimed.

## Por que eu fiz isso?
Trabalhei no CHAA (Complexo Hospitalar Astrogildo de Azevedo) e tive a oportunidade de criar uma automação em formato de aplicativo para o download de exames laboratoriais por período de internação. Desenvolvi com interface visual intuitiva para que qualquer profissional do hospital possa utilizar facilmente no dia a dia, sem precisar usar terminal.

**O que ele faz na prática:**
- Loga automaticamente no sistema do laboratório via navegador (Selenium WebDriver), reconectando sozinho se a sessão expirar.
- Filtra todas as Ordens de Serviço (OS) do paciente pelos intervalos de data e hora de cada internação.
- Baixa os laudos em PDF de forma automatizada e fecha o modal de visualização a cada exame.
- Lê o conteúdo textual dos PDFs baixados para identificar os exames (ex: `LAB1 - HEMOGRAMA_CREATININA.pdf`) e gera também um arquivo consolidado (`LAUDOS_UNIFICADOS.pdf`) por internação.
- Possui máscara automática de data/hora (`DD/MM/AAAA HH:MM`), seletor de pasta de destino e opção de salvar credenciais localmente no `.env`.

## Tecnologias
- **Linguagem:** Python 3.12+
- **Scraping / Automação Web:** Selenium WebDriver + WebDriver Manager
- **Interface Gráfica (GUI):** CustomTkinter
- **Processamento de PDF:** pypdf (identificação de exames e unificação de laudos)
- **Segurança e Configuração:** python-dotenv
- **Empacotamento:** PyInstaller (geração de executável `.exe` standalone)

## Estrutura do Projeto

```text
├── .env.example          # Template para configurar credenciais e pasta padrão
├── requirements.txt      # Dependências diretas do projeto (UTF-8)
├── __main__.py           # Ponto de entrada principal da aplicação
├── app.py                # Interface gráfica em CustomTkinter (thread-safe)
├── bot.py                # Robô Selenium (login, sessão, download e pop-ups)
├── config.py             # Seletores do portal e leitura/gravação do .env
├── utilidades.py         # Pastas, máscara de datas, leitura e fusão de PDFs
├── scripts/
│   └── build_exe.py      # Script para gerar o executável LabimedDownload.exe
└── tests/                # Testes unitários automatizados (unittest)
```

## Como Executar

1. Crie e ative o ambiente virtual:
   ```powershell
   py -m venv venv
   .\venv\Scripts\activate
   ```
2. Instale as dependências:
   ```powershell
   pip install -r requirements.txt
   ```
3. (Opcional) Copie `.env.example` para `.env` ou preencha seu usuário e senha diretamente na tela do aplicativo.
4. Inicie a aplicação:
   ```powershell
   python __main__.py
   ```

## Como Gerar o Executável (.exe)

Para gerar um arquivo único `dist\LabimedDownload.exe` pronto para rodar em computadores sem Python instalado:

```powershell
.\venv\Scripts\python.exe scripts\build_exe.py
```