# VBA Code Review

Denne gjennomgangen tar utgangspunkt i UserForm- og modul-koden du delte.

## Kritiske feil

1. **`Nz`-kall uten tilgjengelig implementasjon.** Både i `ResolvePostsWithPigg` og i den dupliserte koden nederst i modulen brukes `Nz(antRorRad, 0)`, men Excel-VBA har ingen innebygd `Nz`-funksjon. Dersom prosjektet ikke refererer til et egendefinert `Nz`, vil kompilering feile med «Sub eller Function ikke definert». Bytt til en trygg konvertering, for eksempel `IIf(IsNumeric(antRorRad), antRorRad, 0)` eller en egen liten hjelpefunksjon.

    ```vb
    Or (LCase$(overligger) = "rør" And (InStr(overTxt, "rør") > 0 Or InStr(overTxt, "ror") > 0) And antRor = Nz(antRorRad, 0)) _
    Or (LCase$(overligger) = "ingen" And overTxt = "" And Nz(antRorRad, 0) = 0) Then
    ```

2. **Duplisert kodefragment etter `SettInnArbeid`.** Etter `End Sub` følger et sett løse `Dim`-, `For`- og `If`-setninger samt et ekstra `End Sub` uten tilhørende prosedyredeklarasjon. Dette gir «Ugyldig utenfor prosedyre» og hindrer kompilering. Slett blokken helt, eller flytt den inn i en egen prosedyre.

    ```vb
    End Sub




        Dim lastR As Long, r As Long
        Dim ralTxt As String: ralTxt = IIf(lakkert, "ja", "nei")
        ' ...
        ' Hvis ikke funnet, behold verdier fra HentStolpeRegel
    End Sub
    ```

## Funksjonelle forbedringer

1. **Bruk eksplisitt ark i stedet for `ActiveSheet`.** `Set wsCalc = ActiveSheet` gjør at kalkylen skrives til det arket brukeren tilfeldigvis har aktivt. Vurder å peke på kalkylesiden eksplisitt med `ThisWorkbook.Sheets("Kalkyle")` eller lignende for å unngå feilregistrering.
2. **Manglende bruk av dekningsgrad i `SettInnArbeid`.** Prosedyren tar inn `Dekningsgrad`, men den brukes ikke i prisberegningen (kolonne H settes til 0). Dersom montørtimer skal følge samme påslag som øvrige produkter, kan du regne tilsvarende margin eller eksplisitt kommentere hvorfor det ikke skal være påslag.
3. **Valider brukerinput før konvertering.** Flere felter konverteres direkte med `CLng`/`CDbl`. Hvis feltene er tomme, vil det gi runtime-feil. Du kan legge inn enkel validering og feilmeldinger før regningen kjøres.
4. **Sortering av høyder i `UserForm_Initialize`.** Dictionary-iterasjonen returnerer høydene i uspesifisert rekkefølge. Dersom du ønsker stigende sortering i komboboksen kan du legge dem i en `Collection` eller bruke `WorksheetFunction.Sort` før du legger dem til.

## Andre observasjoner

- `ResolvePostsWithPigg` har samme logikk to ganger (én gang riktig innkapslet, én gang i den løse blokken). Når du rydder bort duplikatet, blir modulen lettere å vedlikeholde.
- `SettInnProdukt` setter formel `=D*r*H*r` som tekst. Hvis brukerne har lokal Excel-konfigurasjon med komma som desimal eller semikolon som formelskille, vurder `FormulaLocal` eller å skrive verdien via `Cells(r, 9).Value = wsCalc.Cells(r, 4).Value * wsCalc.Cells(r, 8).Value` for å unngå kulturavhengige feil.

Gi gjerne lyd hvis du ønsker hjelp til konkrete endringer eller forbedringsforslag!
