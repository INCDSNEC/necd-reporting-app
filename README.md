# NECD Reporting (Art. 9)

Aplicație web statică pentru importul fișierelor originale de laborator și introducerea manuală a valorilor, cu verificări tehnice neblocante, afișarea progresului și exportul raportului NECD în XLSX.

## Acces și date

Aplicația se publică gratuit pe GitHub Pages. Codul nu include șablonul oficial, date de laborator sau tokenuri. Datele sunt salvate într-un repository separat, **privat**, `INCDSNEC/necd-reporting-data`.

În aplicație, introdu un GitHub fine-grained personal access token cu drept `Contents: Read and write` limitat la repository-ul de date. Folosește aplicația numai pe dispozitive de încredere. Tokenul este păstrat în `sessionStorage` doar pentru sesiunea browserului, dar orice aplicație web ce primește un token are acces la date în limita acelui token. Nu-l publica și nu-l trimite altor persoane.

Pentru configurare și limite, consultă documentația din pachetul de livrare.
