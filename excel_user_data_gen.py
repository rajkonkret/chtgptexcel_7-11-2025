import pandas as pd

# Generowanie przykładowych danych użytkowników (30 rekordów)
data = {
    "id": range(1, 31),
    "imie": [
        "Jan", "Anna", "Piotr", "Katarzyna", "Marek", "Agnieszka", "Tomasz", "Ewa", "Paweł", "Magdalena",
        "Krzysztof", "Joanna", "Andrzej", "Monika", "Michał", "Barbara", "Rafał", "Aleksandra", "Adam", "Beata",
        "Łukasz", "Dorota", "Grzegorz", "Natalia", "Dariusz", "Karolina", "Sebastian", "Justyna", "Marcin", "Elżbieta"
    ],
    "nazwisko": [
        "Kowalski", "Nowak", "Wiśniewski", "Wójcik", "Kamiński", "Lewandowska", "Zieliński", "Szymańska", "Woźniak", "Dąbrowska",
        "Kozłowski", "Jankowska", "Mazur", "Kwiatkowska", "Krawczyk", "Piotrowska", "Grabowski", "Nowicka", "Pawłowski", "Michalska",
        "Adamczyk", "Zając", "Duda", "Wieczorek", "Jabłoński", "Król", "Majewski", "Olszewska", "Jaworski", "Witkowska"
    ],
}

# Dodanie kolumny email
data["email"] = [
    f"{imie.lower()}.{nazwisko.lower()}@example.com".replace("ł", "l").replace("ś", "s").replace("ń", "n").replace("ż", "z").replace("ź", "z").replace("ć", "c").replace("ó", "o").replace("ą", "a").replace("ę", "e")
    for imie, nazwisko in zip(data["imie"], data["nazwisko"])
]

# Tworzenie DataFrame
df = pd.DataFrame(data)

# Zapis do pliku Excel
excel_path = "uzytkownicy_canvas.xlsx"
df.to_excel(excel_path, index=False)

print(f"Plik zapisany jako {excel_path}")
