# Resume prompt

> Privato, ignorato da git. Prompt di ripresa della sessione: si aggiorna alla fine di ogni sessione con lo stato raggiunto e con un prompt pronto da incollare alla riapertura. Lo stato canonico del progetto resta in `.claude/memory/index.md`; questo file e la comodità per ripartire in fretta, non una seconda fonte di verità.

> In coda a questo file vive l'impronta di ripresa, un commento che non si vede nel rendering e che non si scrive a mano: la registra `python tools/verifica-ripresa.py --registra` come ultimo atto della sessione, dopo i commit. Serve alla sessione successiva per sapere se questo file descriva ancora il presente, e la sua assenza non è un difetto ma la condizione che la verifica riconosce.

<!-- Solo negli alberi di lavoro che non sono quello autorevole (regola alberi-di-lavoro.md): l'avviso seguente va subito sotto il titolo, prima di ogni altra cosa, togliendo questo commento; nell'albero autorevole si toglie tutto il blocco.

> ATTENZIONE: questo albero è sulla branch <branch>, che non è la più avanti. La memoria di progetto valida è in <percorso assoluto dell'albero autorevole>/.claude/memory/: si legge da lì per percorso assoluto, non si copia e non si fonde qui.
-->

## Stato raggiunto

- Data: <YYYY-MM-DD>
- Branch / commit: <branch> / <hash>
- Dove siamo: <una o due frasi sul punto raggiunto>
- Prossimo passo: <azione concreta da cui ripartire>

## Da incollare a Claude alla riapertura

```
Riprendi il progetto <nome progetto>. Procedi cosi, senza leggere tutto:
1. Esegui la skill riprendi: verifica con tools/verifica-ripresa.py che questo file
   descriva lo stato reale, e riportami che cosa risulta non scritto da una sessione
   caduta prima di leggere altro.
2. Leggi .claude/memory/index.md (snapshot: branch, commit, stato schede, punto di ripresa).
3. Leggi .claude/context/current-work.md per la feature attiva.
4. Esegui la skill sync-context per misurare il drift delle schede rispetto a HEAD.
5. Dammi un recap conciso (dove siamo, cosa risulta fatto, cosa la verifica ha trovato di
   non scritto, prossimo passo) e fermati.
Vincoli: niente operazioni git (le faccio io); nessun valore segreto nei file tracciati.
```
