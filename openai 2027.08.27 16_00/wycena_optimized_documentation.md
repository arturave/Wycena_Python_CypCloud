# wycena_optimized.py — dokumentacja

## Cel
Szybkie i stabilne przetwarzanie plików XLSX (w trybie `read_only=True`, `data_only=True`),
opcjonalne cache'owanie cennika oraz generowanie dokumentu WZ (DOCX).
Zawiera minimalistyczne GUI (Tkinter) z obsługą błędów (okna dialogowe).

## Architektura
- `Item`, `AnalysisResult` — proste struktury danych (dataclasses).
- `parse_items_from_xlsx(path)` — parsuje kolumny **Lp, Symbol, Nazwa, Ilość** po nagłówkach.
- `analyze_folder(folder)` — agreguje pozycje z wielu XLSX i raportuje sumy.
- `load_price_list(price_path, file_hash)` — ładuje i cache'uje cennik (CSV/XLSX) z `lru_cache`.
- `generate_wz_doc(...)` — szybkie generowanie WZ w DOCX (bez cen).
- `App` — prosta aplikacja z wyborem folderu i cennika, komunikaty o błędach.

## Jakość kodu
- Adnotacje typów (PEP 484), docstringi Google/NumPy-friendly.
- Konfiguracja **Black** i **Ruff** w `pyproject.toml`.
- Logowanie (logging) dla istotnych kroków.

## Obsługa błędów
- Brak XLSX → `messagebox.showerror` i przerwanie.
- Nieznalezione nagłówki → `messagebox.showerror`.
- Problemy z cennikiem → `messagebox.showwarning`.

## Szybkie uruchomienie
```bash
python wycena_optimized.py
```
W GUI wybierz folder z XLSX i ewentualnie cennik. Kliknij „Analizuj i wygeneruj WZ”.
