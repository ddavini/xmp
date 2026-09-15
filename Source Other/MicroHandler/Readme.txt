Microweb Server http://www.zolnetwork.com/x-mad/
By X-MaD x-mad@zolnetwork.com
Ver 0.9.25 Beta (Vesione Sperimentale)

Microweb server e' un mini web server programmato in VB6
Per farlo girare e' necessario NTSVC.OCX e MSWINSCK.OCX + le VB Runtime ovviamente.

Dalla Versione 0.9.25 e' stato aggiunto il supporto al CGI del PHP. Build In. Per
Usufruirne e' ovviamente necessario tale CGI.

Note per la distribuzione in formato compilato:
Usate pure Microweb senza limitazioni di sorta.
Unica cosa che chiedo e che mi segnalaste gli eventuli Bug.

Note per la distribuzione in formato sorgente:
L'utilizzo e' completamente libero solo due cose dovreste fare.
1- Se utilizzate tutto o in parte il codice per fare un MicroWeb Server evitate di 
   chiamarlo *Microweb* :)
2- Se apportate migliorie al codice siete pregati di spedire tali milgiorie in 
   formato soregente all'autore.
3- Se trovate dei bug sia che li correggiate o meno siete pregati di segralarli all'
   autore. Se poi li correggete anche mandate il Tips in formato sorgente all'autore 
   che e' sempre quello di cui sopra :).

Si ringrazia Roberto Negro per il suo HTTP Virtual Server demo project (C) 1999
Da cui ha tratto spunto Microweb server

Io l'ho fatto per usarlo (come tutti i miei programmi), se poi vien utile pure a voi
la cosa non puo' che rendermi felice.

Note:
- per utilizzare il Progetto in debug asteriscare la Chiamata alla Sub NTSvc Oppure
usare come command line -debug

- La presente vesione di Microweb supporta il solo comando GET

- Microweb e stato testato solo su sistemi con tecnologia NT, in quanto l'autore non
crede nell'esistenza di Windows9x/ME/XP e io non sono riuscito a convincerlo del contrario.

- Nell'attuale stato di sviluppo (Beta Sperimentali) non e' stato dato alcun peso alle 
implicazioni di sicurezza. Ogni critica o suggerimento e' comunque molto gradito.
Anche se qualcosina in piu', sotto questo punto di vista, e' stato aggiunto dalla 
versione 0.9.35:
Temp folder - Per fa si di poter settare i permessi in scrittura solo qui e sui log
Deny String - Stringhe proibite nelle gichieste URL Def("..","%")
IP Ban Chk  - IP a cui non e' permessa la connessione

Ciao,
Me