/**
 * Garażownia — odbiór projektów z planera + krótkie kody projektów
 *
 * CO ROBI
 *   1. Zapisuje każdy przysłany projekt w Arkuszu Google i wysyła powiadomienie na maila.
 *   2. Obsługuje krótkie kody: strona zapisuje układ pod 7-znakowym kodem, a potem
 *      pobiera go z powrotem. Kody starsze niż 48 godzin są kasowane.
 *
 * JAK WDROŻYĆ
 *   1. W Arkuszu: Rozszerzenia -> Apps Script.
 *   2. Skasuj stary kod, wklej ten plik w całości.
 *   3. Zapisz (Ctrl+S).
 *   4. Wdróż -> Zarządzaj wdrożeniami -> ołówek -> Wersja: NOWA WERSJA -> Wdróż.
 *      (Samo zapisanie nic nie zmienia — adres /exec serwuje zamrożoną kopię.)
 *
 * Adres /exec zostaje ten sam, nic nie trzeba zmieniać na stronie.
 */

const MOJ_MAIL = 'patryk@garazownia.com';
const NAZWA_ARKUSZA = 'Projekty';
const NAZWA_KODOW = 'Kody';
const NAZWA_RUCHU = 'Ruch';
const WAZNOSC_KODU_H = 48;

/* ============================ WEJŚCIE ============================ */

function doPost(e) {
  try {
    const dane = JSON.parse(e.postData.contents);

    if (dane.akcja === 'zapisz_kod') {
      zapiszKod(dane);
      return odpowiedz({ ok: true });
    }

    if (dane.akcja === 'wejscie' || dane.akcja === 'zdarzenie') {
      zapiszRuch(dane);
      return odpowiedz({ ok: true });
    }

    zapiszDoArkusza(dane);
    wyslijPowiadomienie(dane);
    wyslijKopieKlientowi(dane);
    return odpowiedz({ ok: true });
  } catch (err) {
    console.error(err);
    return odpowiedz({ ok: false, blad: String(err) });
  }
}

/**
 * Bez parametrów: sprawdzenie, czy wdrożenie żyje.
 * Z parametrem ?kod=XXXXXXX: zwraca zapisany projekt.
 * Parametr ?callback=nazwa opakowuje odpowiedź w JSONP — strona pobiera dane
 * przez znacznik <script>, bo Apps Script nie wystawia nagłówków CORS.
 */
function doGet(e) {
  const kod = e && e.parameter ? e.parameter.kod : null;
  const cb = e && e.parameter ? e.parameter.callback : null;

  if (!kod) {
    return odpowiedz({ ok: true, info: 'Odbiornik projektów Garażowni działa.' }, cb);
  }
  try {
    return odpowiedz(wczytajKod(kod), cb);
  } catch (err) {
    console.error(err);
    return odpowiedz({ ok: false, blad: 'serwer' }, cb);
  }
}

function odpowiedz(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback && /^[A-Za-z_][A-Za-z0-9_]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ======================= KRÓTKIE KODY (48 h) ======================= */

function arkuszKodow() {
  const plik = SpreadsheetApp.getActiveSpreadsheet();
  let ark = plik.getSheetByName(NAZWA_KODOW);
  if (!ark) {
    ark = plik.insertSheet(NAZWA_KODOW);
    ark.appendRow(['Kod', 'Zapisano', 'Wygasa', 'Wymiary', 'Elementów', 'Kod projektu']);
    ark.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1C1F22').setFontColor('#FFFFFF');
    ark.setFrozenRows(1);
    ark.setColumnWidth(6, 420);
  }
  return ark;
}

function zapiszKod(d) {
  const kod = String(d.kod || '').toUpperCase().replace(/[^0-9A-Z]/g, '');
  if (kod.length !== 7) throw new Error('Nieprawidłowy kod');

  const ark = arkuszKodow();
  usunPrzeterminowane(ark);

  const teraz = new Date();
  const wygasa = new Date(teraz.getTime() + WAZNOSC_KODU_H * 3600 * 1000);
  ark.appendRow([
    kod, teraz, wygasa,
    d.wymiary || '', d.liczba_elementow || '', d.kod_projektu || ''
  ]);
}

function wczytajKod(kodWejscie) {
  const kod = String(kodWejscie).toUpperCase().replace(/[^0-9A-Z]/g, '');
  const ark = arkuszKodow();
  const dane = ark.getDataRange().getValues();
  const teraz = new Date();

  for (let i = dane.length - 1; i >= 1; i--) {   // od najnowszych
    if (String(dane[i][0]).toUpperCase() !== kod) continue;
    const wygasa = new Date(dane[i][2]);
    if (teraz > wygasa) return { ok: false, blad: 'wygasl' };
    return { ok: true, kod_projektu: String(dane[i][5]) };
  }
  return { ok: false, blad: 'brak' };
}

/** Kasuje wiersze starsze niż 48 h. Wołane przy każdym zapisie, więc arkusz się nie rozrasta. */
function usunPrzeterminowane(ark) {
  const dane = ark.getDataRange().getValues();
  const teraz = new Date();
  for (let i = dane.length - 1; i >= 1; i--) {
    const wygasa = new Date(dane[i][2]);
    if (isNaN(wygasa.getTime()) || teraz > wygasa) ark.deleteRow(i + 1);
  }
}

/* ===================== WŁASNY LICZNIK RUCHU =====================
   GoatCounter potrafi paść (503 na całej usłudze, 20.08.2026), a wtedy nie
   wiadomo, ile osób przyszło z posta. Zapisujemy więc wejścia u siebie:
   data, kampania z linku (?utm_campaign=...), domena źródłowa, rodzaj
   urządzenia. Bez ciasteczek, bez adresu IP — nie da się z tego wskazać osoby. */

function arkuszRuchu() {
  const plik = SpreadsheetApp.getActiveSpreadsheet();
  let ark = plik.getSheetByName(NAZWA_RUCHU);
  if (!ark) {
    ark = plik.insertSheet(NAZWA_RUCHU);
    ark.appendRow(['Data', 'Rodzaj', 'Kampania', 'Skad', 'Urzadzenie', 'Ekran']);
    ark.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1C1F22').setFontColor('#FFFFFF');
    ark.setFrozenRows(1);
    ark.setColumnWidth(1, 150);
    ark.setColumnWidth(2, 190);
  }
  return ark;
}

function zapiszRuch(d) {
  const ark = arkuszRuchu();
  const rodzaj = d.akcja === 'wejscie' ? 'wejscie' : ('etap: ' + String(d.nazwa || '').slice(0, 60));
  ark.appendRow([
    new Date(),
    rodzaj,
    String(d.kampania || '').slice(0, 60),
    String(d.skad || '').slice(0, 80),
    String(d.urzadzenie || ''),
    String(d.ekran || '')
  ]);
}

/* ===================== PROJEKTY WYSŁANE MAILEM ===================== */

function zapiszDoArkusza(d) {
  const plik = SpreadsheetApp.getActiveSpreadsheet();
  let ark = plik.getSheetByName(NAZWA_ARKUSZA);

  if (!ark) {
    ark = plik.insertSheet(NAZWA_ARKUSZA);
    ark.appendRow([
      'Data', 'E-mail klienta', 'Wymiary garażu', 'Liczba elementów',
      'Wycena', 'Podsumowanie', 'Kod projektu'
    ]);
    ark.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1C1F22').setFontColor('#FFFFFF');
    ark.setFrozenRows(1);
    ark.setColumnWidth(6, 420);
    ark.setColumnWidth(7, 260);
  }

  ark.appendRow([
    new Date(),
    d.email || '',
    d.wymiary || '',
    d.liczba_elementow || '',
    d.wycena || '',
    d.podsumowanie || '',
    d.kod_projektu || ''
  ]);
}


/* Kopia dla klienta — obiecujemy na stronie, że dostanie listę i kod projektu,
   więc musi to przyjść samo, a nie dopiero gdy Patryk usiądzie do maila. */
function wyslijKopieKlientowi(d) {
  const mail = String(d.email || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(mail)) return;
  MailApp.sendEmail({
    to: mail,
    subject: 'Twój projekt garażu ' + (d.wymiary || ''),
    replyTo: MOJ_MAIL,
    body:
      'Dziękuję za przesłanie projektu. Poniżej wszystko, co ustawiłeś w planerze.\n\n' +
      (d.podsumowanie || '') +
      '\n\nKod projektu: ' + (d.kod_projektu || '-') + '\n' +
      'Wejdź na https://garazownia.com, kliknij "Wczytaj kod" i wklej te znaki,\n' +
      'żeby wrócić do tego układu na dowolnym urządzeniu. Kod działa 48 godzin.\n\n' +
      'Odpiszę osobiście z uwagami, zwykle w ciągu jednego dnia roboczego.\n\n' +
      'Patryk Szatkowski\n' + MOJ_MAIL
  });
}

function wyslijPowiadomienie(d) {
  const temat = 'Nowy projekt garażu: ' + (d.wymiary || '') +
                ' (' + (d.liczba_elementow || 0) + ' elem.)';

  const tresc =
    'Klient wysłał projekt z planera.\n\n' +
    'KONTAKT: ' + (d.email || '—') + '\n' +
    'WYCENA:  ' + (d.wycena || '—') + '\n' +
    '\n----------------------------------------\n\n' +
    (d.podsumowanie || '') +
    '\n\n----------------------------------------\n' +
    'Aby zobaczyć ten układ w planerze: otwórz stronę, kliknij "Wczytaj kod"\n' +
    'i wklej kod projektu z końca podsumowania.\n';

  MailApp.sendEmail({
    to: MOJ_MAIL,
    subject: temat,
    body: tresc,
    replyTo: d.email || MOJ_MAIL
  });
}

/* =========================== TESTY =========================== */

/** Sprawdza zapis do arkusza i wysyłkę maila. Uruchom z edytora. */
function testowyZapis() {
  const przyklad = {
    email: 'klient@przyklad.pl',
    wymiary: '3.4 x 6 x 2.5 m',
    liczba_elementow: 3,
    wycena: '1279 zł',
    podsumowanie: 'GARAŻ: 3.4 × 6 m\nWYPOSAŻENIE: Regał modułowy x 1, Stół warsztatowy x 1',
    kod_projektu: 'TEST'
  };
  zapiszDoArkusza(przyklad);
  wyslijPowiadomienie(przyklad);
}

/** Sprawdza cały obieg krótkiego kodu: zapis, odczyt i reakcję na nieistniejący kod. */
function testKodu() {
  const kod = 'TEST123';
  zapiszKod({ kod: kod, kod_projektu: 'ABC', wymiary: '3 x 5 x 2.5 m', liczba_elementow: 2 });
  Logger.log('zapisany kod -> %s', JSON.stringify(wczytajKod(kod)));
  Logger.log('nieistniejacy -> %s', JSON.stringify(wczytajKod('ZZZZZZZ')));
}
