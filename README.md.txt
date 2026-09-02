# Bot de Automação para Laudos Médicos

Um script em Python com interface gráfica que fiz para automatizar o download e a organização de laudos de exames laboratoriais.

##  Por que eu fiz isso?
Trabalhei no CHAA (Complexo Astrogildo de Azevedo) e tive a oportunidade de começar a criar uma automação em forma de aplicativo para o download de exames. Fiz em forma de aplicativo utilizando o tkinter do Python para ser mais fácil para as pessoas que não tem tanto domínio com computadores possam usufruir dele.

Fiz esse bot para fazer o trabalho braçal sozinho.

**O que ele faz na prática:**
- Loga automaticamente no sistema do laboratório via navegador (Selenium).
- Baixa os relatórios de forma automatizada.
- Lê e organiza os arquivos em pastas locais padronizadas.
- Tem uma interface visual para o usuário só apertar um botão e acompanhar os logs, sem precisar encostar no terminal.

## Tecnologias
- **Linguagem:** Python
- **Scraping / Automação:** Selenium WebDriver
- **Interface Gráfica (GUI):** Tkinter
- **Segurança:** Python-dotenv (pra não deixar senha vazada no código)

## Estrutura do Código
Organizei o código em módulos pra não virar bagunça:

```text
├── .env.example        # Template para você colocar suas senhas locais
├── requirements.txt    # Dependências do projeto
├── __main__.py         # Arquivo principal que roda a aplicação
├── app.py              # Front-end (Interface gráfica)
├── bot.py              # Back-end (O bot do Selenium que faz a navegação)
├── config.py           # Gerenciamento do arquivo .env
└── utilidades.py       # Funções soltas para manipular arquivos e pastas