# FuturoTCC — Bot de Automação para Laudos Médicos (Labimed / CHAA)

## Produto e stack

- Produto e escopo: Automação desktop com interface gráfica para download e organização por período de internação de laudos laboratoriais em PDF no portal Labimed (Complexo Hospitalar Astrogildo de Azevedo - CHAA).
- Stack: Python 3.12 / 3.14, Tkinter (`ttk`), Selenium WebDriver 4, `webdriver-manager`, `python-dotenv`.
- Diretório local: `C:\Users\migpa\Projetos\futuro-tcc`.

## Comandos canônicos

- Criar ambiente virtual: `py -m venv venv`
- Instalar dependências: `.\venv\Scripts\pip install -r requirements.txt`
- Executar aplicação: `.\venv\Scripts\python __main__.py` (ou `py __main__.py`)
- Verificar sintaxe/compilação: `py -m py_compile __main__.py app.py bot.py config.py utilidades.py`
- Testes automatizados: `py -m unittest discover -s tests` (a configurar conforme plano de implementação)

## Acordos de trabalho

- Leia código e instruções pertinentes antes de editar. Declare hipóteses que alterem comportamento ou escopo.
- Complete o pedido com a menor mudança suficiente. Reutilize padrões, componentes e dependências existentes.
- Preserve alterações locais, validação, tratamento de erros, segurança e requisitos explícitos.
- Não exponha credenciais reais de pacientes ou do sistema hospitalar no código ou no Git (`.env` deve permanecer no `.gitignore`).
- Atualize `RESUMO.md` ao concluir cada etapa relevante.

## Convenções e estrutura atual

- `__main__.py`: ponto de entrada da aplicação Tkinter.
- `app.py`: interface gráfica (`App`), validação de datas de internação e controle de thread de execução.
- `bot.py`: classe `LabimedBot` com automação Selenium (login, verificação de sessão, filtragem por data de OS, download de PDF e fechamento de modal).
- `config.py`: seletores DOM do portal Labimed e leitura de credenciais via variáveis de ambiente (`LABIMED_USUARIO`, `LABIMED_SENHA`).
- `utilidades.py`: manipulação de diretórios temporários/finais, conversão de datas do portal (`DD/MM/AAAA - HH:MM`) e espera de conclusão de download.
