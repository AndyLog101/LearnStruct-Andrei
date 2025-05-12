import PySimpleGUI as sg
import os
import fitz  # PyMuPDF
import requests
import json

OLLAMA_URL = "http://localhost:11434/api/chat"  # endpoint oficial chat Ollama
MODEL_NAME = "mistral"  # sau "llama2", dacă ai descărcat altul

# ------------------ FUNCȚII ------------------

def incarca_text_din_pdf(cale):
    try:
        doc = fitz.open(cale)
        text = ""
        for pagina in doc:
            text += pagina.get_text()
        doc.close()
        return text
    except Exception as e:
        sg.popup_error("Eroare la citirea fișierului PDF:", e)
        return None

def chat_ollama(prompt):
    try:
        response = requests.post(
            "http://localhost:11434/api/chat",
            headers={"Content-Type": "application/json"},
            data=json.dumps({
                "model": MODEL_NAME,
                "stream": False,  # ← AICI E CHEIA!
                "messages": [{"role": "user", "content": prompt}]
            })
        )
        response.raise_for_status()
        result = response.json()
        return result['message']['content']
    except Exception as e:
        sg.popup_error("Eroare la comunicarea cu Ollama:", e)
        return None

def genereaza_rezumat(text):
    prompt = f"Rezumă următorul text în română cu enter după fiecare propoziție:\n\n{text}"
    return chat_ollama(prompt)

def genereaza_test(text, nr_intrebari, nr_raspunsuri, multiple):
    tip = "cu un singur răspuns corect" if not multiple else "cu răspunsuri multiple"
    prompt = (
        f"Generează un test format din {nr_intrebari} întrebări {tip}, fiecare cu "
        f"{nr_raspunsuri} variante de răspuns, pe baza textului următor:\n\n{text}"
    )
    return chat_ollama(prompt)

def salveaza_fisier(nume, continut):
    try:
        with open(nume, "w", encoding="utf-8") as f:
            f.write(continut)
    except Exception as e:
        sg.popup_error(f"Eroare la salvarea fișierului '{nume}':", e)

# ------------------ INTERFAȚĂ ------------------

def create_gui():
    sg.theme("LightBlue")

    layout = [
        [sg.Text("Selectează fișierul PDF:"), sg.Input(), sg.FileBrowse(file_types=(("PDF files", "*.pdf"),))],
        [sg.Text("Număr întrebări:"), sg.InputText("5", size=(5, 1))],
        [sg.Text("Variante de răspuns pe întrebare:"), sg.InputText("4", size=(5, 1))],
        [sg.Checkbox("Răspunsuri multiple corecte", default=False, key="multiple")],
        [sg.Button("Generează Rezumat & Test", key='_generate_'), sg.Button("Ieșire")],
        [sg.Output(size=(100, 20))]
    ]

    window = sg.Window("Aplicație Rezumat + Test (Ollama local)", layout)

    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, "Ieșire"):
            break
        if event == '_generate_':
            cale_fisier = values[0]
            nr_intrebari = values[1]
            nr_raspunsuri = values[2]
            multiple = values["multiple"]

            if not cale_fisier or not os.path.exists(cale_fisier):
                sg.popup_error("Selectează un fișier PDF valid!")
                continue

            try:
                nr_intrebari = int(nr_intrebari)
                nr_raspunsuri = int(nr_raspunsuri)
            except ValueError:
                sg.popup_error("Numărul de întrebări și răspunsuri trebuie să fie numere întregi.")
                continue

            text = incarca_text_din_pdf(cale_fisier)
            if not text:
                continue

            print("Generez rezumat...")
            rezumat = genereaza_rezumat(text)
            if rezumat:
                salveaza_fisier("rezumat.txt", rezumat)
                print("Rezumat salvat în 'rezumat.txt'.")
            else:
                continue

            print("Generez test...")
            test = genereaza_test(text, nr_intrebari, nr_raspunsuri, multiple)
            if test:
                salveaza_fisier("test_generat.txt", test)
                print("Test salvat în 'test_generat.txt'.")
                sg.popup("Rezumatul și testul au fost generate și salvate cu succes!")

    window.close()

# ------------------ EXECUȚIE ------------------

if __name__ == '__main__':
    create_gui()
