/**
 * ============================================================================
 * DE HOOFDFUNCTIE: 'main'
 * Dit is de "aan-knop" van het script. Google roept deze functie elk uur aan.
 * ============================================================================
 */
function main() {
  // We leggen eerst de basisinstellingen vast: waar zijn we en hoe laat is het nu?
  const mijnTijdzone = "Europe/Brussels";
  const mijnLocale = "be";
  const nu = new Date(); // Maak een digitale momentopname van de huidige tijd en datum
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // Open het Excel-bestand waar we in werken

  const sheet = ss.getSheetByName("Logs");

  // We vertellen de spreadsheet dat we in de Belgische tijdzone en landinstelling werken
  ss.setSpreadsheetTimeZone(mijnTijdzone);
  ss.setSpreadsheetLocale(mijnLocale);

  /**
   * STAP 1: DE TIJDCHECK
   * Waarom? Om energie en data te besparen. We hoeven de weerberichten niet 
   * midden in de nacht op te halen als er toch niets verandert voor de planten.
   */
  const huidigUurBelgie = parseInt(Utilities.formatDate(nu, mijnTijdzone, "H"));
  if (huidigUurBelgie < 3) {
    console.log("Het is nacht (" + huidigUurBelgie + "u). Het script gaat weer slapen tot 03:00u.");
    return; // De 'return' is de noodstop: alles hieronder wordt niet meer uitgevoerd.
  }

  /**
   * STAP 2: DE POORTWACHTER (SLIDING WINDOW LOGICA)
   * Waarom? We willen voorkomen dat we elk uur dezelfde data ophalen. 
   * Het script kijkt in de allerlaatste cel van de kolom AH (nummer 34).
   */
  const lastRow = sheet.getLastRow(); // Hoeveel rijen zijn er nu in totaal?
  const vandaagLabel = Utilities.formatDate(nu, mijnTijdzone, "yyyy-MM-dd"); // Vandaag in tekst, bijv. "2026-04-07"

  if (lastRow > 1) { // Als de lijst niet leeg is...
    // Haal de waarde op uit de allerlaatste cel in kolom AH (nummer 34)
    const laatsteUpdateCel = sheet.getRange(lastRow, 34).getValue();

    // Controleren we: is de datum in die cel gelijk aan de datum van vandaag?
    if (laatsteUpdateCel instanceof Date) {
      const laatsteUpdateDatum = Utilities.formatDate(laatsteUpdateCel, mijnTijdzone, "yyyy-MM-dd");

      if (laatsteUpdateDatum === vandaagLabel) {
        // Als de datums gelijk zijn, is de update vandaag al gelukt. We stoppen.
        console.log("De update is vandaag al uitgevoerd. We hoeven niets te doen.");
        return;
      }
    }
  }

  /**
   * STAP 3: DE OPDRACHT GEVEN
   * Als we hier zijn gekomen, is de tijdcheck gepasseerd en heeft de poortwachter 
   * gezien dat er vandaag nog geen update is geweest. Nu roepen we de weer-functie aan.
   */
  console.log("Actie! We gaan de weerdata ophalen bij Make.com...");
  logDetailedWeatherViaMake(sheet, mijnTijdzone, nu);
}

/**
 * ============================================================================
 * WEERLOGICA: logDetailedWeatherViaMake
 * Deze functie gaat "bellen" met het internet om weercijfers te krijgen.
 * ============================================================================
 */
function logDetailedWeatherViaMake(sheet, tz, timestamp) {
  // Coördinaten voor de locatie (Gent)
  const lat = 51.05;
  const lon = 3.73;
  // Het 'telefoonnummer' (URL) van onze automatisering op Make.com
  const makeWebhookUrl = "https://hook.eu1.make.com/j7oqi1y36idxksuunjl24mxmwf1v9yl6";
  const fullUrl = `${makeWebhookUrl}?lat=${lat}&lon=${lon}`;

  // We vragen Make.com om de gegevens. De 'retry' functie zorgt dat we bij een bezette lijn later terugbellen.
  const response = callMakeWithRetry(fullUrl);

  // console.log("HTTP status code: " + response.getResponseCode());
  // console.log("JSON payload");
  // console.log(response.getContentText());

  const res = JSON.parse(response.getContentText()); // De ruwe tekst omzetten in een digitale mappenstructuur

  const daily = res.daily;   // De map met daggegevens
  const hourly = res.hourly; // De map met uurgegevens
  const hourlyPollen = res.hourlyPollen; // De map met uurgegevens van de pollen data

  /**
   * HULPMIDDEL: getAvgForDay
   * Een klein rekenmachientje om het gemiddelde van 24 losse uren te berekenen.
   */
  const getAvgForDay = (arr, d) => {
    // Check of de lijst überhaupt bestaat en niet leeg is
    if (!arr || !Array.isArray(arr) || arr.length === 0) {
      // console.warn("Waarschuwing: Geen data gevonden voor dag " + d);
      return 0; // Geef een 0 terug als de data ontbreekt, zo crasht de rest niet
    }

    const start = d * 24;
    const daySlice = arr.slice(start, start + 24);

    if (daySlice.length === 0) return 0;

    const sum = daySlice.reduce((a, b) => a + b, 0);
    return sum / daySlice.length;
  };

  const getMaxForDay = (arr, d) => {
    // Controleer of de lijst bestaat
    if (!arr || !Array.isArray(arr) || arr.length === 0) {
      return 0;
    }

    const start = d * 24;
    const daySlice = arr.slice(start, start + 24);

    // Math.max kan niet direct een lijst (array) lezen, 
    // daarom gebruiken we de 'spread' operator (...) om de getallen eruit te strooien.
    const max = Math.max(...daySlice);

    // Als de lijst toevallig alleen uit nulls of foutieve data bestaat, 
    // geeft Math.max soms -Infinity. Dit vangen we hier op:
    return max === -Infinity ? 0 : max;
  };

  /**
   * STAP 4: OVERZICHT VAN BESTAANDE DATA
   * We kijken welke datums al in onze spreadsheet staan (kolom A).
   */
  const lastRow = sheet.getLastRow();
  let existingDates = [];
  if (lastRow > 0) {
    // We maken een simpele lijst met tekstuele datums die we al hebben
    existingDates = sheet.getRange(1, 1, lastRow, 1).getValues().map(row =>
      row[0] instanceof Date ? Utilities.formatDate(row[0], tz, "yyyy-MM-dd") : String(row[0])
    );
  }

  // --- NIEUW: VERZAMELBAK VOOR POLLEN DATA ---
  let allePollenData = [];

  /**
   * STAP 5: DE LUS (HET HERHALEN VOOR 4 DAGEN)
   * We gaan nu door een lusje: doe dit voor dag 0 (vandaag), 1, 2 en 3.
   */
  for (let d = 0; d < 4; d++) {
    const datumLabel = daily.time[d]; // De datum van de dag die we nu verwerken
    const startUurIndex = d * 24;

    /**
     * LOGICA: NACHTVORST VOORSPELLING
     * Voor elke dag kijken we 3 nachten vooruit. Als de grondtemperatuur 
     * ergens onder 0 zakt, zetten we een waarschuwing ("VANNACHT", "MORGEN", etc.)
     */
    let forecastVorst = "GEEN";
    const vorstLabels = ["VANNACHT", "MORGEN", "OVERMORGEN"];
    for (let f = 1; f <= 3; f++) {
      const fStart = (d + f) * 24;
      if (hourly.soil_temperature_0cm.length >= fStart + 24) {
        const minGrondForecast = Math.min(...hourly.soil_temperature_0cm.slice(fStart, fStart + 24));
        if (minGrondForecast <= 0) {
          forecastVorst = vorstLabels[f - 1]; // We hebben vorst gevonden!
          break; // Stop met verder zoeken voor deze dag
        }
      }
    }

    /**
     * STAP 6: DE RIJ BOUWEN
     * We maken een lange rij met alle verzamelde cijfers, van kolom A tot kolom AH.
     */
    const row = [
      Utilities.formatDate(new Date(daily.time[d]), tz, "yyyy-MM-dd"), // A: Datum
      Number(daily.temperature_2m_min[d]),                            // B: Laagste luchttemp
      Number(daily.apparent_temperature_min[d]),                      // C: Hoe koud het echt voelt
      Number(daily.temperature_2m_max[d]),                            // D: Hoogste luchttemp
      Number(daily.apparent_temperature_max[d]),                      // E: Hoogste gevoelstemp
      Number(Math.min(...hourly.soil_temperature_0cm.slice(startUurIndex, startUurIndex + 24))), // F: Bevriest de grond?
      Number(daily.precipitation_sum[d]),                               // G: Hoeveelheid regen in mm
      Number(getAvgForDay(hourly.surface_pressure, d)),                // H: Luchtdruk (gemiddelde)
      Number(daily.windspeed_10m_max[d] / 3.6),                        // I: Wind in meter per seconde
      Number(daily.winddirection_10m_dominant[d]),                     // J: Waar komt de wind vandaan?
      Number(getAvgForDay(hourly.relativehumidity_2m, d) / 100),       // K: Vochtigheid van de lucht
      Number(getAvgForDay(hourly.soil_temperature_6cm, d)),            // L: Temperatuur van de aarde
      Number(getAvgForDay(hourly.soil_moisture_3_to_9cm, d)),          // M: Hoe nat is de aarde?
      Number(daily.daylight_duration[d] / 3600),                       // N: Aantal uren zonlicht
      Number(getAvgForDay(hourly.cloudcover, d) / 100),                // O: Hoe bewolkt is het?
      Number(getAvgForDay(hourly.cloudcover_high, d) / 100),           // P: Hoge bewolking
      Number(getAvgForDay(hourly.cloudcover_low, d) / 100),            // Q: Lage bewolking
      Number(getAvgForDay(hourly.freezinglevel_height, d)),            // R: Hoe hoog in de lucht vriest het?
      daily.temperature_2m_min[d] < 2 ? "JA" : "NEE",                  // S: Directe vorstwaarschuwing
      Number(daily.apparent_temperature_max[d]),                       // T: Hitte index
      (daily.temperature_2m_max[d] > 10 && daily.precipitation_sum[d] < 2) ? "GUNSTIG" : "MATIG/SLECHT", // U: Plantconditie
      (daily.temperature_2m_max[d] > 18 && daily.precipitation_sum[d] === 0 && daily.windspeed_10m_max[d] < 20) ? "UITSTEKEND" : "NIET IDEAAL", // V: Groei-index
      Number(daily.et0_fao_evapotranspiration[d]),                      // W: Hoeveel water verdampt er?
      Utilities.parseDate(daily.sunrise[d], tz, "yyyy-MM-dd'T'HH:mm"),  // X: Tijdstip zonsopgang
      Utilities.parseDate(daily.sunset[d], tz, "yyyy-MM-dd'T'HH:mm"),   // Y: Tijdstip zonsondergang
      Number(daily.uv_index_max[d]),                                    // Z: Kracht van de zon (UV)
      forecastVorst,                                                    // AA: Onze nachtvorst voorspelling
      Number(getMaxForDay(hourlyPollen.grass_pollen, d)),               // AB: Gras
      Number(getMaxForDay(hourlyPollen.alder_pollen, d)),               // AC: Els
      Number(getMaxForDay(hourlyPollen.ragweed_pollen, d)),             // AD: Ambrosia (Ragweed)
      Number(getMaxForDay(hourlyPollen.olive_pollen, d)),               // AE: Es (Olijffamilie)
      Number(getMaxForDay(hourlyPollen.birch_pollen, d)),               // AF: Berk
      Number(getMaxForDay(hourlyPollen.mugwort_pollen, d)),             // AG: Bijvoet
      timestamp                                                         // AH: De 'stempel' van wanneer dit gelogd is
    ];

    /**
     * STAP: E-MAIL NOTIFICATIE (Alleen voor vandaag: d=0)
     */
    if (d === 0) {
      stuurDroogteEmailViaRowArray(row);
    }

    // --- NIEUW: DATA OPSLAAN VOOR MEERDAAGSE CHECK ---
    allePollenData.push(row);

    /**
     * STAP 7: UPDATEN OF TOEVOEGEN (UPSERT)
     * Bestaat de datum al in onze lijst? Dan overschrijven we de oude info met de nieuwste.
     * Is het een nieuwe dag? Dan plakken we die onderaan de lijst.
     */
    const rowIndex = existingDates.indexOf(datumLabel);
    let doelRij; // Wordt gebruikt om later de focus te zetten

    if (rowIndex !== -1) {
      // De datum is gevonden! Overschrijf de rij op die plek.
      doelRij = rowIndex + 1;
      sheet.getRange(doelRij, 1, 1, row.length).setValues([row]);
      console.log("Gegevens vernieuwd voor: " + datumLabel);
    } else {
      // De datum is nieuw: voeg een nieuwe regel toe onderaan de spreadsheet.
      sheet.appendRow(row);
      doelRij = sheet.getLastRow();
      console.log("Nieuwe dag toegevoegd aan de lijst: " + datumLabel);
    }

    // De dag van vandaag selecteren (alleen bij de eerste stap van de lus)
    if (d === 0) {
      sheet.getRange(doelRij, 1).activate();
    }
  } // Einde van de D-lus

  // --- NIEUW: DE POLLENCHECK VOOR DE KOMENDE 4 DAGEN ---
  console.log("Gecapteerde pollen data : " + allePollenData);
  checkPollenVoorMeerdereDagen(allePollenData);

  /**
   * STAP 8: DE AFWERKING
   * Nu de cijfers in de sheet staan, maken we het mooi.
   */
  applyFormatting(sheet); // Zorg voor de juiste komma's en procenttekens
  applyConditionalFormatting(sheet, tz); // Geef cellen een kleurtje bij gevaar of kansen

  console.log("Het script is succesvol afgerond.");
}

/**
 * ============================================================================
 * E-MAIL ALARM FUNCTIE
 * ============================================================================
 */
function stuurUitdorgingsAlarm(emailAdres, verdamping, temp, reden) {
  const onderwerp = "⚠️ Aardbei-Alarm: De piramide heeft dorst!";
  const bericht = "Dag!\n\nJe automatische plantenwachter heeft een risico op uitdroging gedetecteerd:\n\n" +
    "- Reden: " + reden + "\n" +
    "- Voorspelde verdamping: " + verdamping.toFixed(2) + " mm\n" +
    "- Maximum temperatuur: " + temp.toFixed(1) + " °C\n\n" +
    "Vergeet niet om de aardbei-piramide vandaag extra water te geven!\n\n" +
    "Groetjes,\nJe Google Script";

  MailApp.sendEmail(emailAdres, onderwerp, bericht);
  console.log("Alarm e-mail verzonden naar: " + emailAdres);
}

/**
 * ============================================================================
 * RETRY LOGIC: callMakeWithRetry
 * Waarom? Internetverbindingen kunnen soms haperen. Als Make.com te druk is, 
 * krijgt het script een foutmelding (Code 500). In plaats van op te geven, 
 * wacht dit script even en probeert het daarna opnieuw (maximaal 4 keer).
 * ============================================================================
 */
function callMakeWithRetry(url) {
  const delays = [15000, 30000, 60000, 120000]; // Wachttijden in milliseconden
  const maxAttempts = 4; // Maximaal 4 pogingen

  for (let i = 0; i < maxAttempts; i++) {
    try {
      console.log("Verbinding maken... Poging " + (i + 1));
      const response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true, "method": "get" });
      const responseCode = response.getResponseCode();

      // Code 200 betekent: "Gelukt! Ik heb de data voor je."
      if (responseCode === 200) {
        return response;
      }

      console.warn("Lijn bezet (Code " + responseCode + "). We proberen het zo opnieuw.");
    } catch (e) {
      console.error("Netwerkfout: " + e.message);
    }

    // Als het niet de laatste poging is, even pauzeren ('slapen')
    if (i < maxAttempts - 1) {
      Utilities.sleep(delays[i]);
    }
  }
  // Als we hier komen, zijn alle 4 de pogingen mislukt.
  throw new Error("Het lukt niet om verbinding te maken met Make.com. Controleer de automatisering.");
}

/**
 * ============================================================================
 * VISUELE FUNCTIES: Opmaak & Kleuren
 * Deze functies veranderen niets aan de data, alleen aan hoe het eruit ziet.
 * ============================================================================
 */
function applyFormatting(sheet) {
  // We vertellen Google Sheets welk soort getal in welke kolom staat
  sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd");  // Datums
  sheet.getRange("B:W").setNumberFormat("#,##0.00");    // Getallen met 2 cijfers na de komma
  sheet.getRange("M:M").setNumberFormat("#,##0.000");   // bodemvocht
  sheet.getRange("K:K").setNumberFormat("0%");          // Procenten voor vochtigheid
  sheet.getRange("O:Q").setNumberFormat("0%");          // Procenten voor bewolking
  sheet.getRange("X:Y").setNumberFormat("HH:mm");       // Kloktijden
  sheet.getRange("AB:AG").setNumberFormat("0");         // Pollen
  sheet.getRange("AH:AH").setNumberFormat("yyyy-mm-dd HH:mm"); // Tijdstempel van de run
}

function applyConditionalFormatting(sheet, mijnTijdzone) {
  sheet.clearConditionalFormatRules();
  const rules = [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return;

  const bereikHeleSheet = sheet.getRange(2, 1, lastRow - 1, lastCol);

  // --- ALGEMENE OPMAAK ---

  // VANDAAG (Rij-breed): Vetgedrukt en blauwe achtergrond
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$A2=VANDAAG()")
    .setBold(true)
    .setBackground("#E8F0FE")
    .setRanges([bereikHeleSheet])
    .build());

  // --- SPECIFIEKE KOLOMMEN ---

  // F: Grondtemp (<0 Rood, <5 Oranje)
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground("#F8D7DA").setRanges([sheet.getRange("F2:F")]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(5).setBackground("#FFF3CD").setRanges([sheet.getRange("F2:F")]).build());

  // S: Frost Warning
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("JA").setBackground("#FAD7DA").setRanges([sheet.getRange("S2:S")]).build());

  // AA: Forecast (Vorst vannacht of morgen)
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied("=OR($AA2=\"VANNACHT\"; $AA2=\"MORGEN\")").setBackground("#F8D7DA").setRanges([sheet.getRange("AA2:AA")]).build());

  // L: Bodemtemp (<5 Blauw, >10 Groen)
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(5).setBackground("#D1ECF1").setRanges([sheet.getRange("L2:L")]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(10).setBackground("#D4EDDA").setRanges([sheet.getRange("L2:L")]).build());

  // M: Bodemvocht (<0.15 Geel/Droog, >0.35 Blauw/Nat)
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.15).setBackground("#FFF3CD").setRanges([sheet.getRange("M2:M")]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0.35).setBackground("#CCE5FF").setRanges([sheet.getRange("M2:M")]).build());

  // Z: UV Index (>= 5 Oranje)
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(5).setBackground("#FFE5D0").setRanges([sheet.getRange("Z2:Z")]).build());

  // O: Bewolking (>= 80% Grijs)
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.8).setBackground("#E2E3E5").setRanges([sheet.getRange("O2:O")]).build());

  // --- POLLEN FORMATTERING (AB t/m AG) ---

  // 1. GRAS POLLEN (Kolom AB) - Zeer allergeen
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(1).setBackground("#FFF2CC").setRanges([sheet.getRange("AB2:AB")]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(10).setBackground("#FCE5CD").setRanges([sheet.getRange("AB2:AB")]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(50).setBackground("#F8D7DA").setRanges([sheet.getRange("AB2:AB")]).build());

  // 2. ELS & BIJVOET (Kolom AC & AG) - Matig allergeen
  const matigBereik = [sheet.getRange("AC2:AC"), sheet.getRange("AG2:AG")];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(30).setBackground("#FFF2CC").setRanges(matigBereik).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(80).setBackground("#FCE5CD").setRanges(matigBereik).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(150).setBackground("#F8D7DA").setRanges(matigBereik).build());

  // 3. AMBROSIA (Kolom AD) - Extreem allergeen (Ragweed)
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(1).setBackground("#FFF2CC").setRanges([sheet.getRange("AD2:AD")]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(5).setBackground("#FCE5CD").setRanges([sheet.getRange("AD2:AD")]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(20).setBackground("#F8D7DA").setRanges([sheet.getRange("AD2:AD")]).build());

  // 4. ES / OLIJF (Kolom AE) - Sterk allergeen
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(10).setBackground("#FFF2CC").setRanges([sheet.getRange("AE2:AE")]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(40).setBackground("#FCE5CD").setRanges([sheet.getRange("AE2:AE")]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(100).setBackground("#F8D7DA").setRanges([sheet.getRange("AE2:AE")]).build());

  // 5. BERK (Kolom AF) - Sterk allergeen
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(10).setBackground("#FFF2CC").setRanges([sheet.getRange("AF2:AF")]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(50).setBackground("#FCE5CD").setRanges([sheet.getRange("AF2:AF")]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(100).setBackground("#F8D7DA").setRanges([sheet.getRange("AF2:AF")]).build());

  // Sla alle regels op
  sheet.setConditionalFormatRules(rules);

  // Teken de dikke zwarte randen rond vandaag
  highlightTodayRowWithBorders(sheet, mijnTijdzone);
}

function highlightTodayRowWithBorders(sheet, timezone) {
  const nu = new Date();
  const vandaagLabel = Utilities.formatDate(nu, timezone, "yyyy-MM-dd");

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return;

  // Maak schoon
  sheet.getRange(2, 1, lastRow, lastCol).setBorder(false, false, false, false, false, false);

  const data = sheet.getRange(1, 1, lastRow, 1).getValues();

  for (let i = 0; i < data.length; i++) {
    let celWaarde = data[i][0];
    let datumTekst = "";

    if (celWaarde instanceof Date) {
      datumTekst = Utilities.formatDate(celWaarde, timezone, "yyyy-MM-dd");
    } else if (typeof celWaarde === 'string') {
      datumTekst = celWaarde.substring(0, 10);
    }

    if (datumTekst === vandaagLabel) {
      const rijIndex = i + 1;
      // Teken de dikke randen
      sheet.getRange(rijIndex, 1, 1, lastCol)
        .setBorder(true, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_THICK);
      return;
    }
  }
}

function stuurDroogteEmailViaRowArray(row) {
  const dochterEmail = "vankets.margot@gmail.com";
  const papaEmail = "bert@vankets.com";
  const et0Vandaag = row[22];
  const tempMaxVandaag = row[3];

  let reden = "";
  if (et0Vandaag > 3.5) {
    reden = "De verdamping is vandaag erg hoog (" + et0Vandaag.toFixed(2) + " mm).";
  } else if (tempMaxVandaag > 25) {
    reden = "De temperatuur stijgt boven de 25°C, terracotta potten drogen nu snel uit.";
  }

  if (reden !== "") {
    stuurUitdorgingsAlarm(dochterEmail, et0Vandaag, tempMaxVandaag, reden);
    stuurUitdorgingsAlarm(papaEmail, et0Vandaag, tempMaxVandaag, reden);
  }
}

// function checkPollenViaRowArray(row) {
//   // Deze functie laten we staan voor compatibiliteit, maar wordt niet meer aangeroepen in 'main'
//   const ontvangers = [
//     { telefoon: "32478385741", apiKey: "3716704" },
//     { telefoon: "32468581422", apiKey: "8510233" }
//   ];

//   const pollenMapping = [
//     { naam: "Grassen",  index: 27, type: "gras" },
//     { naam: "Elzen",    index: 28, type: "boom" },
//     { naam: "Ambrosia", index: 29, type: "kruid" },
//     { naam: "Olijf/Es", index: 30, type: "boom" },
//     { naam: "Berken",   index: 31, type: "boom" },
//     { naam: "Bijvoet",  index: 32, type: "kruid" }
//   ];

//   let meldingen = [];

//   pollenMapping.forEach(item => {
//     const waarde = row[item.index];
//     if (typeof waarde === 'number') {
//       const info = bepaalPollenStatus(waarde, item.type);
//       if (info.score >= 3) {
//         meldingen.push(`- *${item.naam}*: ${waarde.toFixed(1)} gr/m³\n  👉 Status: ${info.label} (${info.score}/5)`);
//       }
//     }
//   });

//   if (meldingen.length > 0) {
//     const bericht = "⚠️ *Pollen Waarschuwing* ⚠️\n\n" +
//                     "De volgende soorten zijn vandaag verhoogd aanwezig:\n\n" +
//                     meldingen.join("\n\n") + 
//                     "\n\n_Bron: Uw plantenlogboek_";
    
//     ontvangers.forEach(p => stuurWhatsApp(p.telefoon, p.apiKey, bericht));
//   }
// }

/**
 * Hulpmiddel om de ernst te bepalen op basis van soort en concentratie
 */
function bepaalPollenStatus(waarde, type) {
  let score = 1;
  let label = "Zeer laag";

  // Drempels bepalen op basis van type
  let drempels = [];
  if (type === "gras")  { drempels = [5, 20, 50, 150]; }
  else if (type === "boom") { drempels = [10, 25, 150, 500]; }
  else { drempels = [5, 20, 50, 150]; } // Kruid/Ambrosia

  // Score berekenen
  if (waarde >= drempels[3]) { score = 5; label = "🔴 *ZEER HOOG*"; }
  else if (waarde >= drempels[2]) { score = 4; label = "🟠 Hoog"; }
  else if (waarde >= drempels[1]) { score = 3; label = "🟡 Matig"; }
  else if (waarde >= drempels[0]) { score = 2; label = "🟢 Laag"; }
  
  return { score: score, label: label };
}

/**
 * NIEUWE FUNCTIE: checkPollenVoorMeerdereDagen
 * Deze verwerkt de array van 4 dagen en stuurt één gebundeld WhatsApp bericht.
 */
function checkPollenVoorMeerdereDagen(alleDagenRows) {
  const ontvangers = [
    { telefoon: "32478385741", apiKey: "3716704" },
    { telefoon: "32468581422", apiKey: "8510233" }
  ];

  const pollenMapping = [
    { naam: "Grassen",  index: 27, type: "gras" },
    { naam: "Elzen",    index: 28, type: "boom" },
    { naam: "Ambrosia", index: 29, type: "kruid" },
    { naam: "Olijf/Es", index: 30, type: "boom" },
    { naam: "Berken",   index: 31, type: "boom" },
    { naam: "Bijvoet",  index: 32, type: "kruid" }
  ];

  let totaalBericht = "🗓️ *POLLENVOORSPELLING* 🗓️\n\n";
  let heeftRelevanteData = false;

  alleDagenRows.forEach((row, i) => {
    const datum = row[0]; // De datum van de dag
    let dagMeldingen = [];

    pollenMapping.forEach(item => {
      const waarde = row[item.index];
      if (typeof waarde === 'number') {
        console.log("Trigger check:", item.naam, waarde, info.score);
        const info = bepaalPollenStatus(waarde, item.type);
        // Alleen tonen als score matig (3) of hoger is
        if (info.score >= 3) {
          dagMeldingen.push(`- ${item.naam}: ${info.label}`);
          heeftRelevanteData = true;
          console.log("Relevante pollen data gevonden voor " + item.naam);
        }
      }
    });

    // Label bepalen (Vandaag vs de datum)
    const dagLabel = (i === 0) ? "*VANDAAG*" : `*${datum}*`;
    
    if (dagMeldingen.length > 0) {
      totaalBericht += `${dagLabel}:\n${dagMeldingen.join("\n")}\n\n`;
    } else {
      totaalBericht += `${dagLabel}:\n 🟢 Lage concentraties\n\n`;
    }
  });

  // Verstuur bericht alleen als er ergens in de 4 dagen relevante pollen zijn
  if (heeftRelevanteData) {
    console.log("Relevante pollen data gevonden. Message zal verstuurd worden via Whatsapp");
    totaalBericht += "_Bron: Uw plantenlogboek_";
    ontvangers.forEach(p => stuurWhatsApp(p.telefoon, p.apiKey, totaalBericht));
  }
}

function stuurWhatsApp(telefoon, apiKey, bericht) {
  // De URL moet 'encoded' zijn zodat spaties en emoticons goed aankomen
  const url = "https://api.callmebot.com/whatsapp.php?phone=" + telefoon +
    "&text=" + encodeURIComponent(bericht) +
    "&apikey=" + apiKey;

  try {
    const response = UrlFetchApp.fetch(url);
    if (response.getResponseCode() == 200) {
      console.log("WhatsApp bericht succesvol verzonden!");
    }
  } catch (e) {
    console.error("Fout bij verzenden WhatsApp: " + e.message);
  }
}
