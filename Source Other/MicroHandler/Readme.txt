Microweb Server http://www.zolnetwork.com/x-mad/
By X-MaD x-mad@zolnetwork.com
Ver 0.9.25 Beta (Vesione Sperimentale)

(EN) Microweb Server http://www.zolnetwork.com/x-mad/
By X-MaD x-mad@zolnetwork.com
Ver 0.9.25 Beta (Experimental Version)

Microweb server e' un mini web server programmato in VB6
Per farlo girare e' necessario NTSVC.OCX e MSWINSCK.OCX + le VB Runtime ovviamente.

(EN) Microweb server is a mini web server written in VB6.
To run it you need NTSVC.OCX and MSWINSCK.OCX, plus the VB Runtime, obviously.

Dalla Versione 0.9.25 e' stato aggiunto il supporto al CGI del PHP. Build In. Per
Usufruirne e' ovviamente necessario tale CGI.

(EN) As of version 0.9.25, built-in support for PHP CGI has been added. To use
it you obviously need that CGI installed.

Note per la distribuzione in formato compilato:
Usate pure Microweb senza limitazioni di sorta.
Unica cosa che chiedo e che mi segnalaste gli eventuli Bug.

(EN) Notes for distribution in compiled form:
Feel free to use Microweb without any restrictions.
The only thing I ask is that you report any bugs you find.

Note per la distribuzione in formato sorgente:
L'utilizzo e' completamente libero solo due cose dovreste fare.
1- Se utilizzate tutto o in parte il codice per fare un MicroWeb Server evitate di
   chiamarlo *Microweb* :)
2- Se apportate migliorie al codice siete pregati di spedire tali milgiorie in
   formato soregente all'autore.
3- Se trovate dei bug sia che li correggiate o meno siete pregati di segralarli all'
   autore. Se poi li correggete anche mandate il Tips in formato sorgente all'autore
   che e' sempre quello di cui sopra :).

(EN) Notes for distribution in source form:
Use is completely free, you're just asked to do two things.
1- If you use all or part of the code to make a MicroWeb Server, please avoid
   calling it *Microweb* :)
2- If you make improvements to the code, please send those improvements in
   source form to the author.
3- If you find bugs, whether you fix them or not, please report them to the
   author. If you do fix them, please also send the fix in source form to the
   author, who is always the same one mentioned above :).

Si ringrazia Roberto Negro per il suo HTTP Virtual Server demo project (C) 1999
Da cui ha tratto spunto Microweb server

(EN) Thanks to Roberto Negro for his HTTP Virtual Server demo project (C) 1999,
which Microweb server took inspiration from.

Io l'ho fatto per usarlo (come tutti i miei programmi), se poi vien utile pure a voi
la cosa non puo' che rendermi felice.

(EN) I made it to use it myself (like all my programs); if it turns out to be
useful to you too, that can only make me happy.

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

(EN) Notes:
- to run the Project in debug mode, comment out (put an asterisk before) the
  call to the NTSvc Sub, or use -debug as the command line

- This version of Microweb only supports the GET command

- Microweb has only been tested on systems with NT technology, since the author
  doesn't believe in the existence of Windows9x/ME/XP and I haven't managed to
  convince him otherwise.

- In the current state of development (Experimental Beta), no attention has
  been given to security implications. Any criticism or suggestion is
  nonetheless very welcome. Although a bit more was added on that front
  starting from version 0.9.35:
  Temp folder - To allow setting write permissions only here and on the logs
  Deny String - Strings forbidden in URL requests, e.g. ("..", "%")
  IP Ban Chk  - IPs that are not allowed to connect

Ciao,
Me

(EN) Bye,
Me
