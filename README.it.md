# XML2Excel

[![Italiano](https://img.shields.io/badge/lang-it-green.svg)](README.it.md)
[![English](https://img.shields.io/badge/lang-en-red.svg)](README.md)
[![Python](https://img.shields.io/badge/python-3.6%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-yellow.svg)](LICENSE)

[English Version](README.md)

Questo script Python converte file XML in formato Excel, creando un foglio di calcolo separato per ogni elemento di secondo livello (figli diretti dell'elemento radice).

## Funzionalità

- Converte automaticamente file XML in fogli di calcolo Excel
- Crea un foglio Excel separato per ogni elemento di secondo livello nell'XML
- Genera automaticamente intestazioni di colonna basate sugli attributi name dei tag column
- Mantiene i valori corretti per ogni record
- Supporta una struttura XML gerarchica con record multipli per sezione

## Requisiti

- Python 3.6+
- Librerie: pandas, openpyxl, xml.etree.ElementTree
- Per installare le dipendenze necessarie:

```
pip install pandas openpyxl
```

- In alternativa, è possibile installare tutte le dipendenze dal file requirements.txt:

```
pip install -r requirements.txt
```

## Utilizzo

- Posizionare il file XML e lo script Python nella stessa directory
- Eseguire lo script:

```
python script.py
```

- Inserire il nome del file XML quando richiesto
- Il file Excel verrà generato nella stessa directory con lo stesso nome del file XML ma con estensione .xlsx

## Struttura XML supportata

Lo script è progettato per lavorare con una struttura XML come questa:

<root>
    <sezione1>
        <record>
            <column type='tipo_dati' name='nome_colonna1'>valore1</column>
            <column type='tipo_dati' name='nome_colonna2'>valore2</column>
            <!-- Altri campi -->
        </record>
        <record>
            <column type='tipo_dati' name='nome_colonna1'>valore3</column>
            <column type='tipo_dati' name='nome_colonna2'>valore4</column>
            <!-- Altri campi -->
        </record>
    </sezione1>
    <sezione2>
        <record>
            <column type='tipo_dati' name='nome_colonna3'>valore5</column>
            <column type='tipo_dati' name='nome_colonna4'>valore6</column>
            <!-- Altri campi -->
        </record>
        <!-- Altri record -->
    </sezione2>
    <!-- Altre sezioni -->
</root>

## Output

- Ogni foglio conterrà tutti i record corrispondenti alle sezioni XML
- In ogni foglio le colonne saranno quelle definite dall'attributo name nei tag column

## Note

- I nomi dei fogli Excel sono limitati a 31 caratteri (limitazione di Excel)
- In caso di nomi di fogli duplicati, verrà aggiunto un numero progressivo
- Se un campo è vuoto nel file XML, la corrispondente cella nel foglio Excel sarà vuota
- L'attributo type nei tag column viene ignorato durante la conversione

## Risoluzione problemi

Se lo script genera un errore:

- Assicurati che il file XML sia nella stessa directory dello script
- Verifica che il file XML sia ben formato e segua la struttura prevista
- Controlla di aver installato tutte le librerie richieste
- Verifica di avere i permessi per scrivere nella directory corrente