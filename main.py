import pandas as pd
import re
pd.options.mode.chained_assignment = None


metryczka_naglowki = [
['Kobieta', 'Mężczyzna'],
['Poniżej 18', '18 – 24', '25 – 35', '36 – 45', '46 – 55', 'Powyżej 55'],
['Wieś', 'Miasto do 10 tys.', 'Miasto powyżej 10 tys. do 50 tys.', 'Miasto powyżej 50 tys. do 100 tys.', 'Miasto powyżej 100 tys. do 250 tys.', 'Miasto powyżej 250 tys.'],
['Podstawowe', 'Gimnazjalne', 'Zawodowe', 'Średnie', 'Wyższe'],
['Brak', 'Otwarta próba przewodnikowska', 'Przewodnik/Przewodniczka', 'Podharcmistrz/Podharcmistrzyni', 'Harcmistrz/Harcmistrzyni'],
['Białostocka', 'Dolnośląska', 'Gdańska', 'Kielecka', 'Krakowska', 'Kujawsko-Pomorska', 'Lubelska', 'Łódzka', 'Mazowiecka', 'Opolska', 'Podkarpacka', 'Stołeczna', 'Śląska', 'Warmińsko-Mazurska', 'Wielkopolska', 'Zachodniopomorska', 'Ziemi Lubuskiej']]
    

def tworzenie_formatow(excel):
    pytanie_format = excel.add_format({
    'font_color': 'green',
    'font_size': 16,
    })

    kolumna_format = excel.add_format({
     'fg_color': 'green',
     'font_color': 'white',
     'align': 'center'
    })

    procent_format = excel.add_format({
    'font_color': 'green'
    })

    kilka_format = excel.add_format({
    'bold': True
    })
    return [pytanie_format, kolumna_format, procent_format, kilka_format]


def _sort_odp(list):
    return sorted(list, reverse=True)


def tytul(sheet, title, row, i, pytanie_format):
    #Rysuje tytuł bazy danych lub pytania
    if i == 0:
        sheet.write(row, 0, title, pytanie_format)
    else:
        title = title.iloc[:, 0].name
        if len(title) < 30:
            name = f'2.{i} ' + title.split('. ')[1]
        else:
            name = re.findall(r"\. .+[\.\?:]", title)[0]
            name = f'1.{i} ' + name.split('. ')[1] 
        sheet.write(row, 0, name, pytanie_format)
    return row+2


def _metryczka_tytul(name):
    return name.split(' ')[1]


def _wyciaganie_odpowiedzi(kolumna):
    #Wyciąga z nazwy pytania wielokrotnego wyboru odpowiedź
    odp = re.findall(r'\[[\w\s() – ./]+\]', kolumna)
    odp[1] = odp[1].replace('[', '')
    odp[1] = odp[1].replace(']', '')
    return odp[1]


#Jak będę mieć czas, to wywalić to
def _tytuly_kolumn(sheet, pytanie, row, kolumna_format, mul):
    #Rysuje numer i nazwę pytania
    #Rysuje tytuły kolumn oraz sumę odpowiedzi
    sheet.merge_range(row, 0, row+1, 0, ' ', kolumna_format)
    sheet.merge_range(row, 1, row+1, 1, 'Razem:', kolumna_format)
    sheet.merge_range(row, 2, row+1, 2, ' ', kolumna_format)
    sub = 1
    if len(pytanie.columns) == 1:
        pytanie = pytanie.iloc[:, 0]
        odp_sorted = _sort_odp(pytanie.unique())
        if mul == 100:
            sub = len(pytanie.index)
        for id, odp in enumerate(odp_sorted):
            sheet.write(row+2+id, 0, odp)
            sheet.write(row+2+id, 1, (pytanie.loc[pytanie == odp].count() / sub * mul // 1))
        sheet.write(row + 2 + len(pytanie.unique()), 0, 'Razem:')
        sheet.write(row + 2 + len(pytanie.unique()), 1, len(pytanie.index))
    else:
        for id, kolumna in enumerate(pytanie.columns):
            series = pytanie[kolumna]
            odp = _wyciaganie_odpowiedzi(kolumna)
            sheet.write(row+2+id, 0, odp)
            if odp == 'Inne':
                sheet.write(row+2+id, 1, series.count())
            else:
                sheet.write(row+2+id, 1, (series == 'Tak').sum())
        sheet.write(row + 2 + len(pytanie.columns), 0, 'Razem:')
        sheet.write(row + 2 + len(pytanie.columns), 1, len(pytanie.iloc[:, 0].index))
    return row


def _rysowanie_jedno(sheet, pytanie, metryczka, row, kolumna_format, norm, mul):
    #Rysuje tabelę pytań jednokrotnego wyboru
    pytanie.dropna(inplace=True)
    row = _tytuly_kolumn(sheet, pd.DataFrame(pytanie), row, kolumna_format, mul)
    odp_sorted = _sort_odp(pytanie.unique())
    col = 3
    for id, nazwa_columny in enumerate(metryczka.columns):
        cross_tab = pd.crosstab(pytanie, metryczka[nazwa_columny], normalize=norm) * mul // 1
        sheet.merge_range(row, col, row, col+len(metryczka_naglowki[id])-1, _metryczka_tytul(nazwa_columny), kolumna_format)
        for c, el in enumerate(metryczka_naglowki[id]):
            sheet.write(row+1, col+c, el, kolumna_format)
            for odp_id, odp in enumerate(odp_sorted):
                try:
                    num = cross_tab[el][odp]
                except Exception:
                    num = ' '
                if num == 0:
                    num = ' '
                sheet.write(row+2+odp_id, col+c, num)
            sheet.write(row + len(cross_tab) + 2, col+c, len(pytanie[metryczka[nazwa_columny] == el].index))
        col += len(metryczka_naglowki[id]) + 1
        sheet.merge_range(row, col-1, row+1, col-1, ' ', kolumna_format)
    row += len(pytanie.unique()) + 4
    return row


def jednokrotne(sheet, pytanie, metryczka, row, i, formaty):
    pytanie_format = formaty[0]
    kolumna_format = formaty[1]
    procent_format = formaty[2]
    kilka_format = formaty[3]
    #Korzystając z innych funkcji rysuje całe pytanie jednokrotnego wyboru
    row = tytul(sheet, pd.DataFrame(pytanie), row, i, pytanie_format)
    sheet.write(row-1, 0, 'Liczba odpowiedzi', procent_format)
    row = _rysowanie_jedno(sheet, pytanie, metryczka, row, kolumna_format, False, 1)
    if len(pytanie.unique()) <= 20:
        sheet.write(row-1, 0, 'Procent odpowiedzi', procent_format)
        row = _rysowanie_jedno(sheet, pytanie, metryczka, row, kolumna_format, 'columns', 100)
    return row


def _rysowanie_wielo(sheet, pytania, metryczka, row, kolumna_format, norm, mul):
    #Razem z tytul i _tytuly_kolumn rysuje pytania wielokrotnego wyboru
    row = _tytuly_kolumn(sheet, pytania, row, kolumna_format, mul)
    col = 3
    sub = 1
    for id, nazwa_kolumny in enumerate(metryczka.columns):
        sheet.merge_range(row, col, row, col+len(metryczka_naglowki[id])-1, _metryczka_tytul(nazwa_kolumny), kolumna_format)
        m = metryczka[nazwa_kolumny]
        for c, el in enumerate(metryczka_naglowki[id]):
            sheet.write(row+1, col+c, el, kolumna_format)
            if norm == 'columns':
                sub = len(pytania[m == el].index)
                if sub == 0:
                    sub = 1
            for df_id in range(len(pytania.columns)):
                pytanie = pytania.iloc[:, df_id]
                odp = _wyciaganie_odpowiedzi(pytanie.name)
                if odp == 'Inne':
                    num = pytanie[m == el].count() / sub * mul // 1
                else:
                    temp = pytanie == 'Tak'
                    num = pytanie[temp & (m == el)].count() / sub * mul // 1
                if num == 0:
                    num = ''
                sheet.write(row+2+df_id, col+c, num)
            sheet.write(row + len(pytania.columns) + 2, col+c, len(pytania[m == el].index))
        col += len(metryczka_naglowki[id]) + 1
        sheet.merge_range(row, col-1, row+1, col-1, ' ', kolumna_format)
    row += len(pytania.columns) + 4
    return row


def _wielo_krzyzowe(sheet, pytania, row, kolumna_format):
    sheet.write(row, 0, '', kolumna_format)
    for row_id, wiersz in enumerate(pytania.columns):
        odp = _wyciaganie_odpowiedzi(wiersz)
        sheet.write(row + row_id + 1, 0, odp)
        for col_id, kolumna in enumerate(pytania.columns):
            odp = _wyciaganie_odpowiedzi(kolumna)
            sheet.write(row, 1 + col_id, odp, kolumna_format)
            df = pd.concat([pytania[wiersz], pytania[kolumna]], axis=1)
            df = df[df[wiersz] != 'Nie']
            df = df[df[kolumna] != 'Nie']
            df = df.dropna()
            num = len(df.index)
            if num == 0:
                num = ' '
            sheet.write(row + row_id + 1, 1 + col_id, num)
        sheet.write(row + len(pytania.columns) + 1, 1 + row_id, pytania[wiersz].count() - (pytania[wiersz] == 'Nie').sum())
    row += len(pytania.columns) + 1
    sheet.write(row, 0, 'Razem:')
    row += 2
    return row


def wielokrotne(sheet, pytania, metryczka, row, i, formaty):
    #Razem z tytul i _tytuly_kolumn rysuje pytania wielokrotnego wyboru
    pytanie_format = formaty[0]
    kolumna_format = formaty[1]
    procent_format = formaty[2]
    kilka_format = formaty[3]
    row = tytul(sheet, pytania, row, i, pytanie_format)
    sheet.write(row-1, 0, 'Liczba odpowiedzi', procent_format)
    row = _rysowanie_wielo(sheet, pytania, metryczka, row, kolumna_format, False, 1)
    sheet.write(row-1, 0, 'Procent odpowiedzi', procent_format)
    row = _rysowanie_wielo(sheet, pytania, metryczka, row, kolumna_format, 'columns', 100)
    sheet.write(row-1, 0, 'Tablica krzyżowa - liczba odpowiedzi', procent_format)
    row = _wielo_krzyzowe(sheet, pytania, row, kolumna_format)
    sheet.write(row-1, 0, 'Inne - liczba odpowiedzi', procent_format)
    row = _rysowanie_jedno(sheet, pytania.iloc[:, -1], metryczka, row, kolumna_format, False, 1)
    sheet.write(row-1, 0, 'Inne - procent odpowiedzi', procent_format)
    row = _rysowanie_jedno(sheet, pytania.iloc[:, -1], metryczka, row, kolumna_format, 'columns', 100)
    return row


def _rysowanie_jedno_kilka(sheet, pytania, metryczka, row, kolumna_format, kilka_format, norm, mul):
    #Rysuje tabelę pytań jednokrotnego wyboru
    pytania.dropna(inplace=True)
    # row = _tytuly_kolumn(sheet, pd.DataFrame(pytania.iloc[:, 0]), row, kolumna_format, mul)
    sheet.merge_range(row, 0, row+1, 0, ' ', kolumna_format)
    sheet.merge_range(row, 1, row+1, 1, 'Razem:', kolumna_format)
    sheet.merge_range(row, 2, row+1, 2, ' ', kolumna_format)
    sub = 1
    col = 3
    for id, nazwa_columny in enumerate(metryczka.columns):
        sheet.merge_range(row, col, row, col+len(metryczka_naglowki[id])-1, _metryczka_tytul(nazwa_columny), kolumna_format)
        for c, el in enumerate(metryczka_naglowki[id]):
            sheet.write(row+1, col+c, el, kolumna_format)
        col += len(metryczka_naglowki[id]) + 1
        sheet.merge_range(row, col-1, row+1, col-1, ' ', kolumna_format)
    for pytanie_name in pytania.columns:
        col = 3
        pytanie = pytania[pytanie_name]
        sheet.write(row+2, 0, _wyciaganie_odpowiedzi(pytanie.name), kilka_format)
        row += 1
        odp_sorted = _sort_odp(pytanie.unique())
        if mul == 100:
            sub = len(pytanie.index)
        for id, odp in enumerate(odp_sorted):
            sheet.write(row+2+id, 0, odp)
            sheet.write(row+2+id, 1, (pytanie.loc[pytanie == odp].count() / sub * mul // 1))
        sheet.write(row+len(pytanie.unique())+2, 0, 'Razem: ')
        sheet.write(row+len(pytanie.unique())+2, 1, len(pytanie.index))
        for id, nazwa_columny in enumerate(metryczka.columns):
            cross_tab = pd.crosstab(pytanie, metryczka[nazwa_columny], normalize=norm) * mul // 1
            for c, el in enumerate(metryczka_naglowki[id]):
                for odp_id, odp in enumerate(odp_sorted):
                    try:
                        num = cross_tab[el][odp]
                    except Exception:
                        num = ' '
                    if num == 0:
                        num = ' '
                    sheet.write(row+2+odp_id, col+c, num)
                sheet.write(row + len(cross_tab) + 2, col+c, len(pytanie[metryczka[nazwa_columny] == el].index))
            col += len(metryczka_naglowki[id]) + 1
        row += len(pytanie.unique()) + 1
    row += 3
    return row


def kilka_jednokrotne(sheet, pytania, metryczka, row, i, formaty):
    #Korzystając z innych funkcji rysuje całe pytanie jednokrotnego wyboru
    pytanie_format = formaty[0]
    kolumna_format = formaty[1]
    procent_format = formaty[2]
    kilka_format = formaty[3]
    row = tytul(sheet, pd.DataFrame(pytania.iloc[:, 0]), row, i, pytanie_format)
    sheet.write(row-1, 0, 'Liczba odpowiedzi', procent_format)
    row = _rysowanie_jedno_kilka(sheet, pytania, metryczka, row, kolumna_format, kilka_format, False, 1)
    sheet.write(row-1, 0, 'Procent odpowiedzi', procent_format)
    row = _rysowanie_jedno_kilka(sheet, pytania, metryczka, row, kolumna_format, kilka_format, 'columns', 100)
    return row