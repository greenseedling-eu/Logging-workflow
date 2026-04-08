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
  const sheet = ss.getSheetByName("Logs"); // Zoek het tabblad met de naam "Logs"

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
   * Het script kijkt in de allerlaatste cel van de kolom AB (de tijdstempel-kolom).
   */
  const lastRow = sheet.getLastRow(); // Hoeveel rijen zijn er nu in totaal?
  const vandaagLabel = Utilities.formatDate(nu, mijnTijdzone, "yyyy-MM-dd"); // Vandaag in tekst, bijv. "2026-04-07"

  if (lastRow > 1) { // Als de lijst niet leeg is...
    // Haal de waarde op uit de allerlaatste cel in kolom AB (nummer 28)
    const laatsteUpdateCel = sheet.getRange(lastRow, 28).getValue();

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
  const res = JSON.parse(response.getContentText()); // De ruwe tekst omzetten in een digitale mappenstructuur

  const daily = res.daily;   // De map met daggegevens
  const hourly = res.hourly; // De map met uurgegevens

  /**
   * HULPMIDDEL: getAvgForDay
   * Een klein rekenmachientje om het gemiddelde van 24 losse uren te berekenen.
   */
  const getAvgForDay = (arr, d) => {
    const start = d * 24; // Bereken waar de dag begint in de lijst van uren
    const daySlice = arr.slice(start, start + 24); // Pak een 'hap' van 24 uur uit de lijst
    const sum = daySlice.reduce((a, b) => a + b, 0); // Tel alle waarden bij elkaar op
    return sum / daySlice.length; // Deel door 24 om het gemiddelde te krijgen
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
     * We maken een lange rij met alle verzamelde cijfers, van kolom A tot kolom AB.
     */
    const row = [
      Utilities.formatDate(new Date(daily.time[d]), tz, "yyyy-MM-dd"), // A: Datum
      Number(daily.temperature_2m_min[d]),                            // B: Laagste luchttemp
      Number(daily.apparent_temperature_min[d]),                     // C: Hoe koud het echt voelt
      Number(daily.temperature_2m_max[d]),                            // D: Hoogste luchttemp
      Number(daily.apparent_temperature_max[d]),                     // E: Hoogste gevoelstemp
      Number(Math.min(...hourly.soil_temperature_0cm.slice(startUurIndex, startUurIndex + 24))), // F: Bevriest de grond?
      Number(daily.precipitation_sum[d]),                             // G: Hoeveelheid regen in mm
      Number(getAvgForDay(hourly.surface_pressure, d)),               // H: Luchtdruk (gemiddelde)
      Number(daily.windspeed_10m_max[d] / 3.6),                       // I: Wind in meter per seconde
      Number(daily.winddirection_10m_dominant[d]),                    // J: Waar komt de wind vandaan?
      Number(getAvgForDay(hourly.relativehumidity_2m, d) / 100),      // K: Vochtigheid van de lucht
      Number(getAvgForDay(hourly.soil_temperature_6cm, d)),           // L: Temperatuur van de aarde
      Number(getAvgForDay(hourly.soil_moisture_3_to_9cm, d)),         // M: Hoe nat is de aarde?
      Number(daily.daylight_duration[d] / 3600),                     // N: Aantal uren zonlicht
      Number(getAvgForDay(hourly.cloudcover, d) / 100),               // O: Hoe bewolkt is het?
      Number(getAvgForDay(hourly.cloudcover_high, d) / 100),          // P: Hoge bewolking
      Number(getAvgForDay(hourly.cloudcover_low, d) / 100),           // Q: Lage bewolking
      Number(getAvgForDay(hourly.freezinglevel_height, d)),           // R: Hoe hoog in de lucht vriest het?
      daily.temperature_2m_min[d] < 2 ? "JA" : "NEE",                // S: Directe vorstwaarschuwing
      Number(daily.apparent_temperature_max[d]),                     // T: Hitte index
      (daily.temperature_2m_max[d] > 10 && daily.precipitation_sum[d] < 2) ? "GUNSTIG" : "MATIG/SLECHT", // U: Plantconditie
      (daily.temperature_2m_max[d] > 18 && daily.precipitation_sum[d] === 0 && daily.windspeed_10m_max[d] < 20) ? "UITSTEKEND" : "NIET IDEAAL", // V: Groei-index
      Number(daily.et0_fao_evapotranspiration[d]),                   // W: Hoeveel water verdampt er?
      Utilities.parseDate(daily.sunrise[d], tz, "yyyy-MM-dd'T'HH:mm"),// X: Tijdstip zonsopgang
      Utilities.parseDate(daily.sunset[d], tz, "yyyy-MM-dd'T'HH:mm"), // Y: Tijdstip zonsondergang
      Number(daily.uv_index_max[d]),                                 // Z: Kracht van de zon (UV)
      forecastVorst,                                                 // AA: Onze nachtvorst voorspelling
      timestamp                                                      // AB: De 'stempel' van wanneer dit gelogd is
    ];

    /**
     * STAP: E-MAIL NOTIFICATIE (Alleen voor vandaag: d=0)
     */
    if (d === 0) {
      const dochterEmail = "vankets.margot@gmail.com"; // VERVANG DIT DOOR HET ECHTE ADRES
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
      }
    }

    /**
     * ============================================================================
     * E-MAIL FUNCTIE
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
     * STAP 7: UPDATEN OF TOEVOEGEN (UPSERT)
     * Bestaat de datum al in onze lijst? Dan overschrijven we de oude info met de nieuwste.
     * Is het een nieuwe dag? Dan plakken we die onderaan de lijst.
     */
    const rowIndex = existingDates.indexOf(datumLabel);
    if (rowIndex !== -1) {
      // De datum is gevonden! Overschrijf de rij op die plek.
      sheet.getRange(rowIndex + 1, 1, 1, row.length).setValues([row]);
      console.log("Gegevens vernieuwd voor: " + datumLabel);
    } else {
      // De datum is nieuw: voeg een nieuwe regel toe onderaan de spreadsheet.
      sheet.appendRow(row);
      console.log("Nieuwe dag toegevoegd aan de lijst: " + datumLabel);
    }
  }

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
 * RETRY LOGIC: callMakeWithRetry
 * Waarom? Internetverbindingen kunnen soms haperen. Als Make.com te druk is, 
 * krijgt het script een foutmelding (Code 500). In plaats van op te geven, 
 * wacht dit script even en probeert het daarna opnieuw (maximaal 4 keer).
 * ============================================================================
 */
function callMakeWithRetry(url) {
  const delays = [15000, 30000, 60000, 120000]; // Wachttijden in milliseconden (15sec, 30sec, 1min, 2min)
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
  sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd"); // Datums
  sheet.getRange("B:W").setNumberFormat("#,##0.00");    // Getallen met 2 cijfers na de komma
  sheet.getRange("K:K").setNumberFormat("0%");         // Procenten voor vochtigheid
  sheet.getRange("O:Q").setNumberFormat("0%");         // Procenten voor bewolking
  sheet.getRange("X:Y").setNumberFormat("HH:mm");      // Kloktijden
  sheet.getRange("AB:AB").setNumberFormat("yyyy-mm-dd HH:mm"); // Tijdstempel van de run
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

  // --- SPECIFIEKE KOLOMMEN (Uit je backup) ---

  // F: Grondtemp (<0 Rood, <5 Oranje) - Volgorde is belangrijk!
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

  // HAAL DATA OP VANAF RIJ 1 (om de index exact gelijk te laten lopen met het rijnummer)
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
      // Omdat we bij rij 1 zijn begonnen met inlezen, is de index i exact (rijnummer - 1)
      // Dus de werkelijke rij in Sheets is i + 1
      const rijIndex = i + 1;

      // Teken de dikke randen
      sheet.getRange(rijIndex, 1, 1, lastCol)
        .setBorder(true, null, true, null, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_THICK);

      console.log("Match gevonden! Vandaag (" + vandaagLabel + ") staat op rij: " + rijIndex);
      return;
    }
  }
}
