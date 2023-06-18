# CSV to excel
 
Program ma za zadanie wyciąg odpowiedzi z ankiet z platformy LIME i układać je w przystępnej formie do analizy w excelu. Program robi to za pomocą biblioteki xlsxwriter oraz napisanych funkcji. Funkcje są podzielone ze względu na rodzaje pytań: 
- Pytań jednokrotnego wyboru - najprostrza funkcja, która obsługiwała też pytania, w której wypełniający wpisywał własne odpowiedzi
- Pytań wielokrotnego wyboru - zawiera w sobie wiele elementów, dlatego jest najdłuższa. W środku można znaleźć między innymi tabele krzyżowe
- Powiązanych pytań jednokrotnego wyboru - są to pytania, które są normalnie jednokrotnego wyboru, ale są powiązane z innym pytaniem, dlatego dla łatwiejszej analizy w excelu są one połączone

Uwaga: ze względu na dane osobowe nie dołączam pliku csv, na którym pracowałem, ale można zobaczyć gotowy excel