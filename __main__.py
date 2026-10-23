"""
Ponto de entrada — executa:  python .  (ou python app.py)
"""
import customtkinter as ctk
from app import App


def main() -> None:
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    root = ctk.CTk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
