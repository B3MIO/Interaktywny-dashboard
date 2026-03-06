# Szablon Analityczny VBA – Optymalizacja Danych Procesowych

## 📌 O projekcie
Narzędzie stworzone w ramach pracy inżynierskiej na kierunku Zarządzanie Informacją (Polsko-Japońska Akademia Technik Komputerowych). 
Głównym celem projektu jest automatyzacja przetwarzania danych operacyjnych oraz weryfikacji dokumentacji technicznej w środowisku produkcyjnym. 

Projekt rozwiązuje problem manualnej, czasochłonnej obsługi danych, minimalizując ryzyko błędu ludzkiego i standaryzując proces raportowania dla działów inżynierii i jakości.

## 🛠 Technologie i Architektura
* **Środowisko:** Microsoft Excel 
* **Język programowania:** VBA (Visual Basic for Applications)
* **Główne mechanizmy:**
  * Automatyzacja procesów ETL (Extract, Transform, Load) dla surowych danych wejściowych.
  * Walidacja integralności danych z wykorzystaniem [wpisz np. wyrażeń regularnych / dedykowanych funkcji logicznych].
  * Generowanie zestandaryzowanych raportów końcowych.

## ⚙️ Funkcjonalności
* **Import i czyszczenie danych:** Skrypty automatycznie pobierają surowe dane z systemów [np. klasy ERP] i usuwają zduplikowane lub błędne rekordy.
* **Silnik analityczny:** Dekompozycja i mapowanie danych zgodnie z zapisaną logiką biznesową.
* **Moduł raportowy:** Dynamiczne generowanie widoków podsumowujących kluczowe metryki procesu.

## 🚀 Wyniki wdrożenia i KPI
Testy rozwiązania przeprowadzone w ramach pracy inżynierskiej wykazały:
* Skrócenie czasu potrzebnego na weryfikację i przetworzenie pakietu danych o **[99]%** (z [4] godzin do [1] minuty).
* Zwiększenie spójności i bezbłędności generowanej dokumentacji inżynieryjnej.
* Opracowanie skalowalnej struktury, która pozwala na łatwą adaptację narzędzia do nowych linii lub procesów.

## 📂 Struktura repozytorium
* `SZABLON.xlsm` – Główny plik aplikacji (wymaga włączonej obsługi makr).
* `src/` – Wyeksportowane moduły kodu VBA (`*.bas`) do łatwego przeglądu w przeglądarce.
* `data_sample/` – Zanonimizowany, testowy zestaw danych wejściowych.

## 💻 Instrukcja uruchomienia
1. Pobierz repozytorium na dysk lokalny.
2. Umieść pliki z danymi we wskazanym katalogu docelowym (lub zgodnie z konfiguracją w pliku).
3. Uruchom plik `SZABLON.xlsm` i zezwól na wykonywanie makr.
4. Użyj panelu sterowania w głównym arkuszu, aby zainicjować proces przetwarzania.

---
**Autor:** Szymon Walocha  
